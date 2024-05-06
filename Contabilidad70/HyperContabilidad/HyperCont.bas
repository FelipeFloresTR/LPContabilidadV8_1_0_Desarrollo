Attribute VB_Name = "HyperCont"

Option Explicit

'firma
'Public sFileName As String
' Public Im_Logo As PictureBox
''firma

'Calculadora
Public gCalc As Control
Public gTotCalc As Double

Public gFairware As Integer

'Texto Informe Preliminar
Public Const INFO_PRELIMINAR = "Informe Preliminar"

'Anchos de campos para grillas
Public Const FW_CUENTA = 1200
Public Const FW_NUM = 1200

'Tipo de correlativos comprobante
Public Const TCC_UNICO = 1
Public Const TCC_TIPOCOMP = 2

Public gTipoCorrComp As Integer

'ancho columna de números
Public Const G_VALWIDTH = 1200
Public Const G_DVALWIDTH = 1400

'mensaje para Ley 21.210 Moderniza Legislación Tributaria D.O. 24.02.2020
Public gMsgLey21210 As String

'Periodo de correlativos comprobante
Public Const TCC_MENSUAL = 1
Public Const TCC_ANUAL = 2
Public Const TCC_CONTINUO = 3

Global gPerCorrComp As Integer

'estado de un comprobante por omisión, al crearlo
Global gEstadoNewComp As Integer

'Ver comprobantes nulos en Libro Diario (opción)
Global gCompAnuladoLibDiario As Integer

'Se permite abrir más de un mes al mismo tiempo
Global gAbrirMesesParalelo As Boolean

'Se selecciona Imprimir resumido al generar todo comprobante de centralización
Global gImpResCent As Boolean

'Asignar último día del mes o dia especificado a la fecha del comprobante de centralización
Public Const DTCOMPCENT_CURRDEF = 1
Public Const DTCOMPCENT_LASTDAY = 2
Public Const DTCOMPCENT_DEFDAY = 3

Global gDtCompCent As Integer
Global gDayDtCompCent As Integer


'Activo Fijo: Considerar Mes Completo indistintamente la fecha de inicio de utilización
Public gAFMesCompleto As Boolean

'Indica si el título de un comprobante incluye el tipo del comprebante al imprimirlo
Public gTituloTipoComp As Boolean

'Path de importación/Exportación
Public gImportPath As String
Public gExportPath As String

'Saldo Libro de Caja año anterior
Public gSaldoLibroCajaAnoAnt As Double


'Estado de los meses
Global Const EM_NOEXISTE = 0
Global Const EM_ABIERTO = 1
Global Const EM_CERRADO = 2
Global Const EM_ERRONEO = 3
Global Const MAX_ESTADOMES = EM_ERRONEO

Global gEstadoMes(MAX_ESTADOMES)    '"Abierto", "Cerrado", "Erróneo"

'Estado Impresion Lib Oficial
Public Const EL_IMPRESO = 0      'libro oficial impreso
Public Const EL_ANULADO = 1      'impresión de libro oficial anulada
Public Const MAX_ESTADOLIBIMP = EL_ANULADO
Public gEstadoLibImp(MAX_ESTADOLIBIMP) As String

'tipo doc Honorarios (por ahora 1 sólo)
Public Const TIPODOC_HONOR = 1  'único tipo de doc de retención válido por ahora

'tipo de Retención
Public Const TR_HONORARIOS = 1
Public Const TR_DIETA = 2
Public Const TR_OTRO = 3

Public gTipoRetencion(TR_OTRO) As String


'Libros oficiales
Public Const LIBOF_COMPRAS = 1
Public Const LIBOF_VENTAS = 2
Public Const LIBOF_RETEN = 3
Public Const LIBOF_DIARIO = 4
Public Const LIBOF_MAYOR = 5
Public Const LIBOF_INVBAL = 6
Public Const LIBOF_TRIBUTARIO = 7
Public Const LIBOF_CLASIFICADO = 8
Public Const LIBOF_COMPYSALDOS = 9
Public Const LIBOF_ESTRESCLASIF = 10
Public Const LIBOF_ESTRESCOMP = 11
Public Const LIBOF_ESTRESMENSUAL = 12
Public Const LIBOF_INGEGR = 13

Public Const MAX_LIBOF = LIBOF_INGEGR

Public gLibroOficial(MAX_LIBOF) As String


'largo máximo de un número de documento, ahora que es string
Public Const MAX_NUMDOCLEN = 15

'largo máximo de números asociados a un doc tipo Máquina Registradora
Public Const MAX_NUMDOCMRG = 15

'Cantidad máxima de centros de costo o áreas de negocio por desglose en Estado de Resultado
Public Const MAX_DESGLOESTRESULT = 100

'IVA Irrecuperble
Public Const IVAIRREC_CERO = 0
Public Const IVAIRREC_PARCIAL = 1
Public Const IVAIRREC_TOTAL = 2

'Códigos de Error retornados por funciones
Public Const ERR_VALUTM = -2        'falta valor UTM
Public Const ERR_DEFCUENTA = -3     'falta definir cuenta


'Tipos de Movimientos Activo Fijo
Public Const MOVAF_COMPRA = 1
Public Const MOVAF_VENTA = 2
Public Const MOVAF_BAJA = 3
Public Const MAX_TIPOMOVAF = MOVAF_BAJA

Public gMovActivoFijo(MAX_TIPOMOVAF) As String

' Ordenamiento del Plan de Cuentas
Public Const ORDPLAN_COD = 0
Public Const ORDPLAN_NOM = 1
Public Const ORDPLAN_DESC = 2
Public gFldOrdPlan(ORDPLAN_DESC) As String
Public gFindPlan(ORDPLAN_DESC) As String
Public gOrdPlan As Byte


'Fecha de inicio Traspaso F29 Supermercados
Public gFechaInicioSupermercados As Long
Public gFechaInicioTraspasoSupermercados As Long

'Estado Foliación
Public Const EF_EXISTE = 1
Public Const EF_NOEXISTE = 2

'Niveles cuentas
Public Const MAX_NIVELES = 5
Public Const NIVELINI = 0 'Plan de cuenta

'NIVELES DE CUENTAS
Public Const NIVEL_1 = 1
Public Const NIVEL_2 = 2
Public Const NIVEL_3 = 3
Public Const NIVEL_4 = 4
Public Const NIVEL_5 = 5

'Colores para nievles de informes
Global gColores(MAX_NIVELES) As Long

'algunos colores
Public Const COLOR_VERDEOSCURO = &H8000&        'verde oscuro
Public Const COLOR_AZULOSCURO = &HA60000        'azul oscuro
Public Const COLOR_MORADO = &H800080            'morado
Public Const COLOR_GRIS = &HBEBEBE              'gris
Public Const COLOR_GRISLT = &HD6D6D6              'gris claro
Public Const COLOR_GRISLTLT = &HE1E1E1           'gris claro claro

Public Const COLOR_LIGHTBLUE = &HF9ECDD
Public Const COLOR_VLIGHTBLUE = &HFAF1E7


Public Const COLOR_EDITCELL = &HC6FFFF         'amarillo clarito


'Chequeo
Public Const CHECK_ON = 1
Public Const CHECK_OFF = 0


'Marca de Apertura en una cuenta
Public Const MR_APERTURA = 1

'tipo de relación de una entidad con un documento:

Public Const TRE_EMISOR = 1
Public Const TRE_RECEPTOR = 2
Public Const TRE_OTRO = 3

Public gTipoRelEnt(TRE_OTRO) As String

'Franquicia Tributaria Entidad
Public Const FTE_14A = 1         'Art. 14 A Régimen Semi Integrado
Public Const FTE_14DN3 = 2       'Art. 14 D N° 3 Régimen Pro Pyme  General
Public Const FTE_14DN8 = 3       'Art. 14 D N° 8 Régimen Pro Pyme  Transparente
Public Const FTE_14BN1 = 4       'Art. 14 B N° 1 Renta Efectiva sin Balance
Public Const FTE_RRENTASPRES = 5 'Rentas Presuntas
Public Const FTE_OTRO = 6        'Otro

Public gFranqTribEnt(FTE_OTRO) As String


'Impresión de la columna de detalle de un movimiento (el usuario puede optar entre imprimir la descripción, la entidad, el centro de costo o el área de negocio)
Public Const PRTMOV_DESC = 1
Public Const PRTMOV_ENTIDAD = 2
Public Const PRTMOV_CCOSTO = 3
Public Const PRTMOV_AREANEG = 4

Public Const MAX_PRTMOV = PRTMOV_AREANEG

Public gPrtMovDetOpt As Integer

'Public gCompImpEntidad As Integer

'fecha inicio y término de la depreciación instantánea y DecimaParte
Public gFechaInicioDepInstantanea As Long
Public gFechaTerminoDepInstantanea As Long

'fecha inicio y término de la depreciación decima parte 2
Public gFechaInicioDepDecimaParte2 As Long

'fecha inicio y termino depreciación Ley 21.210
Public gFechaInicioDepLey21210 As Long
Public gFechaTerminoDepLey21210 As Long

'fecha inicio y termino depreciación Ley 21.256
Public gFechaInicioDepLey21256 As Long
Public gFechaTerminoDepLey21256 As Long

'tipos de soscios o propietarios
Public Const MAX_TIPOSOCIO = 11
Public gTipoSocio(MAX_TIPOSOCIO) As String




'Configuración para opciones de empresa 'PS
Public Const OPT_ACTUSADO = &H1     'Actualizar automáticamente en la impresión original folios timbrado
Public Const OPT_NOPRTFECHA = &H2   'No imprimir la fecha en los reportes

'**

'Tipo de dato por los que se puede buscar una cta.
Type Niveles_t
   nNiveles              As Integer
   Inicio(MAX_NIVELES)   As Integer
   Largo(MAX_NIVELES)    As Integer
End Type

Public gNiveles  As Niveles_t
Public gLastNivel  As Integer       'Ultimo nivel de la cuenta
Public gLastNivelIFRS As Integer    'último nivel IFRS
Public gFmtCodigoCta  As String     'Formato del codigo cuenta
Public gFmtCodigoIFRS As String     'Formato del codigo IFRS
Public gNivelesIFRS As Niveles_t    'para plan IFRS

'Deficnición de tipo Cuenta
Type Cuenta_t
   Descripcion    As String
   Codigo         As String
   Nombre         As String
   id             As Long
   Nivel          As Integer
   Tipo           As Integer
   Hijos          As Integer
   IdPadre        As Long
   NivelFather    As Integer
   CodFECU        As String
   CodF22         As Integer
   CodF29         As Integer
   CodF22_14Ter   As Integer
End Type


' Cuentas básicas
Type CuentasBas_t

   'Cuentas asociadas a IVA y Otros Impuestos
   IdCtaIVACred      As Long
   IdCtaIVADeb       As Long
   IdCtaOtrosImpCred As Long
   IdCtaOtrosImpDeb  As Long
   
   'contracuenta pago facturas
   IdCtaPagoFacturas As Long   'pago o egreso
   IdCtaCobFacturas  As Long    'cobranza o ingreso
   
   'cuentas retenciones
   IdCtaImpRet       As Long
   IdCtaNetoHon      As Long
   IdCtaNetoDieta    As Long
   IdCtaImpUnico   As Long
   
   'cuentas de Patrimonio y Resultado Ejercicio
   IdCtaPatrimonio   As Long
   IdCtaResEje       As Long
   
   'Cta de crédito IVA para remante IVA año anterior
   IdCtaCredIVA      As Long
   
   'Cta IVA Irrecuperable
   IdCtaIVAIrrec     As Long
   
   'Cta Retención 3% préstamo
   IdCtaRet3Porc     As Long
   
   'Cta 3% Ret. Centralizacion Remuneraciones
   IdCta3PorcCentraRem     As Long
   
    
   'pipe 2699582
    'Cta PpmObligatorio
   IdCtaPpmObligatorio   As Long
   
   'pipe 2699582
    'Cta PpmVoluntario
   IdCtaPpmVoluntario   As Long
   
   'Feña
   IdCtaOdfActivo     As Long
   IdCtaOdfPasivo     As Long
   'Fin Feña
   
End Type
Public gCtasBas As CuentasBas_t

'Notas definidas por el usuario: artículo 100 y nota especial
Type Nota_t
   TxtNota As String
   IncluirBal As Boolean
   IncluirLib As Boolean
   IncluirInfo As Boolean
End Type

Public gNotaArt100 As Nota_t
Public gNotaEspecial As Nota_t

Public Const C_INCNOTABAL = 1
Public Const C_INCNOTALIB = 2
Public Const C_INCNOTALIBBAL = 3

'Detalle Capital Propio Simplificado
Public Const CPS_PARTICIPACIONES = 1
Public Const CPS_DISMINUCIONES = 2
Public Const CPS_GASTOSRECHAZADOS = 3
Public Const CPS_BASEIMPONIBLE = 4
Public Const CPS_RETDIV = 5
Public Const CPS_AUMENTOSCAP = 6
Public Const CPS_GASTOSRECHNOPAGAN40 = 7
Public Const CPS_INRPROPIOS = 8
Public Const CPS_OTROSAJUSTAUMENTOS = 9
Public Const CPS_OTROSAJUSTDISMIN = 10
Public Const CPS_CAPPROPIOTRIBANOANT = 11
Public Const CPS_REPPERDIDAARRASTRE = 12
Public Const CPS_INRPROPIOSPERDIDAS = 13
Public Const CPS_UTILIDADESPERDIDA = 14
Public Const CPS_INGRESODIFERIDO = 15
Public Const CPS_CTDIMPUTABLEIPE = 16
Public Const CPS_INCENTIVOAHORRO = 17
Public Const CPS_IDPCVOLUNTARIO = 18
Public Const CPS_CREDACTFIJOS = 19
Public Const CPS_CREDPARTICIPACIONES = 20


Public gTipoDetCapPropioSimpl(CPS_CREDPARTICIPACIONES) As String

Public Const CPS_TIPOINFO_GENERAL = 1
Public Const CPS_TIPOINFO_VARANUAL = 2


'Estructuras de datos
Type EntidadHR_t
   Nombre As String        'razón social o nombre completo
   TipoContrib As Integer
   Direccion As String
   NombreCorto As String
   Ciudad As String
   Region As Integer
   Comuna As String
   Tel As String
   Fax As String
   email As String
   DirPostal As String
End Type

Type CCosto_t
   id          As Long
   Codigo      As String
   Descrip     As String
End Type

Type Sucursal_t
   id          As Long
   Codigo      As String
   Descrip     As String
End Type

Type AreaNeg_t
   id       As Long
   Codigo   As String
   Descrip  As String
End Type

Type DetMovim_t
   idDetMov As Long
   idcomp   As Long
   idMov    As Long
   idLib    As Byte
   NDoc     As Long
   Hasta    As String
   
End Type

Type CompTipo_t
   Nombre   As String
   Tipo     As String
   Glosa    As String
   TDebe    As Double
   THaber   As Double
   idTipo   As Integer
   idcomp   As Long
End Type

Type Foliacion_t
   UltImpreso     As Long
   FUltImpreso    As Long
   UltTimbrado    As Long
   FUltTimbrado   As Long
   UltUsado       As Long
   FUltUsado      As Long
   Estado         As Integer
End Type

Public gFoliacion    As Foliacion_t

Type FmtImp_t
   Campo As String
   Formato As String
End Type

Type ResLib_t
   Mes As Integer
   TipoLib As Integer
   TipoDoc As Integer
   ActFijo As Integer
   Giro As Integer
   CountTot As Long
   Exento As Double
   Afecto As Double
   IVA As Double
   CountTotNoGiro As Long
   ExentoNoGiro As Double
   AfectoNoGiro As Double
   IVANoGiro As Double
   OtroImp As Double    'aquí está el impuesto de las retenciones
   CountDTE As Long
   IVADTE As Double
   NetoDTE As Double
   IVAIrrecDTE As Double
   OImpDTE As Double
   TipoReten As Integer
   CountExento As Long
   CountExentoNoGiro As Long
   EsSupermercado As Integer
   CountRetParcial As Long
   NetoRetParcial As Double
   DifIVARetParcial As Double
   IVARetenido As Double      '(parcial o total) este valor sólo se obtiene para desplegarlo en el Resumen de Libros Auxiliares
   CountIVAIrrec As Long
   NetoIVAIrrec As Double     'el neto correspondiente al IVA irrecuperable
   
   CodF29Count As Integer
   CodF29Neto As Integer
   CodF29IVA  As Integer
   CodF29CountNoGiro As Integer
   CodF29NetoNoGiro As Integer
   CodF29IVANoGiro  As Integer
   CodF29IVADTE  As Integer
   CodF29CountDTE  As Integer
   CodF29NetoDTE  As Integer
   CodF29IVAIrrecDTE  As Integer
   CodF29AFCount As Integer
   CodF29AFIVA As Integer
   CodF29ExCount As Integer
   CodF29Exento As Integer
   CodF29ExCountNoGiro As Integer
   CodF29ExentoNoGiro As Integer
   CodF29RetHon As Integer
   CodF29RetDieta As Integer
   CodF29IVARet3ro As Integer
   CodF29CountRetParcial As Integer
   CodF29NetoRetParcial As Integer
   CodF29DifIVARetParcial As Integer
   CodF29CountIVAIrrec As Integer
   CodF29NetoIVAIrrec As Integer
   CodF29CountSuper As Integer
   CodF29IVASuper As Integer
   FacCompraRetParcial As Integer
   IVAIrrec As Integer
End Type

Type ResLibCod_t
   CodF29 As Integer
   Valor As Double
End Type

Type ResOImp_t
   Mes As Integer
   TipoLib As Integer
   CodValLib As Integer
   DescValLib As String
   EsRecuperable As Boolean
   CodF29 As Integer
   CodF29_Adic As Integer
   TipoIVARetenido As Integer
   Valor As Double
End Type


Type VarIniFile_t
   VerDTE As Integer                'Opcionas vista Libro Compras Ventas
   VerExento As Integer             '   "
   VerSucursal As Integer           '   "
   VerNumDocHasta As Integer        '   "
   VerNumInterno As Integer         'opción libro Compra y Ventas
   VerOtrosImp As Integer           '   "
   RepetirGlosa As Integer          '   "
   VerCredArt33 As Integer          'Opciones vista Rep. Activo Fijo:
   VerFVenta As Integer             'Activo Fijo Tributario
   VerFUtiliz As Integer            'Activo Fijo Tributario
   VerValCompraHist As Integer      'Activo Fijo Tributario
   VerTipoDep As Integer            'Activo Fijo Tributario
   VerTipoDepHist As Integer        'Activo Fijo Tributario
   VerPatenteRol As Integer         'Activo Fijo Tributario
   VerNombreProy As Integer         'Activo Fijo Tributario
   VerFechaProy As Integer          'Activo Fijo Tributario
   VerMaqReg As Integer             'Libro Ventas
   VerCantBoletas As Integer        'Libro Ventas
   VerDetOtrosImp As Integer        'Opción libro compras y ventas, sólo para listar, no editar
   VerPropIVA As Integer            'Opción libro compras (proporcionalidad del IVA)
   SelEmprPorRUT As Integer         'Ordenamiento de empresas por RUT en la ventana de selección de empresas
   VerFechaCompra As Integer        'Opciones de vista Rep. Activo Fijo Financiero:
   VerValorInicial As Integer       'Activo Fijo Financiero
   VerPjeAmortizacion As Integer    'Activo Fijo Financiero
   VerFactor As Integer             'Activo Fijo Financiero
   VerValorRazonable As Integer     'Activo Fijo Financiero
   VerRevalorizacion As Integer     'Activo Fijo Financiero
   VerLCajaOper As Integer          'Libro de caja
   VerLCajaDTE As Integer           'Libro de caja
   VerLCajaNombre As Integer        'Libro de caja
   VerLCajaIVAIrrec As Integer      'Libro de caja
   VerLCajaOtrosImp As Integer      'Libro de caja
   VerOtrosIngEgr14TER As Integer   'Auditoria Libros Contables
   VerCodCuenta As Integer          'Auditoria Libros Contables
   VerAreaNeg As Integer            'Auditoria Libros Contables
   VerCCosto As Integer             'Auditoria Libros Contables
   VerGlosaComp As Integer          'Auditoria Libros Contables
   VerLibRetDTE As Integer          'Libro de Retenciones
   VerLibRetSucursal As Integer     'Libro de Reterencioes
   VerRet3Porc As Integer           'Libro de Retenciones
   '2690461
   VerLCajaNotCred As Integer           'Libro de caja
   '2690461
End Type

Public gVarIniFile As VarIniFile_t

Public Const CODF29_IMPUNICO = 48

'selección de documentos que ingresan al informe analítico
Type SelTipoDoc_t
   Default As Boolean                  'indica si no tiene nada cargado y por lo tanto corresponde al default
   SelLib(LIB_OTROS) As Boolean        'indica si el libro(i) está seleccionado
   SelTodos(LIB_OTROS) As Boolean      'indica si para el libro(i) están todos los tipos docs seleccionados
   DocsSel(LIB_OTROS, 100) As Boolean  'docs seleccionados (TipoLib, TipoDoc)
   ConsultaSQL As String               'where correspondiente
End Type

Global gSelTipoDoc As SelTipoDoc_t

Type ImpAdic_t
   TipoLib As Integer
   CodTipoValor As Long
   TipoValor As String
   IdCuenta As Long
   CodCuenta As String
   Cuenta As String
   Tasa As Single
   EsRecuperable As Boolean
End Type

Type LstDoc_t
   IdDoc As Long
   IdDocCuota As Long
   MontoCuota As Double
   NumCuotas As String
End Type

Public gDbFacturacion As String        'ruta para relacionar con Base de Datos de Sistema de Facturación


'2850275
Const SG_PASSW_FAIRPAY = "oP,*/'#2j7h7_$3"


Public DbAccess As Database
Public DbAccess2 As Database


Public lEsLPRemu As Boolean
Public lRemuSQLServer As Boolean


Public lPathlDbRemu As String
Public lConnStr As String

Option Compare Binary

'fin 2850275

'3426794
Public RemIVAAnoAnt As Boolean
'3426794

Public Sub ReadIni()
   Dim Rc As Integer
   Dim Buf As String * 200
     
   'Rc = GetPrivateProfileString("Config", "Calc", "Calc.exe", Buf, 120, gIniFile)
   'gCalcApp = Trim(Left(Buf, Rc))
   
   gCalcApp = "Calc.exe"
   
   'Rc = GetPrivateProfileString("Config", "OrdPlan", "0", Buf, 5, gIniFile)
   'gOrdPlan = Val(Left(Buf, Rc))
   
   gOrdPlan = ORDPLAN_COD
   
   Rc = GetPrivateProfileString("Config", "Fairware", "0", Buf, 5, gCfgFile)
   gFairware = Val(Left(Buf, Rc))
   
   Buf = GetIniString(gIniFile, "Reportes", "PageNumReg", PAGE_NUMREG)
   gPageNumReg = vFmt(Buf)
      
   Buf = GetIniString(gIniFile, "Opciones", "VerDTE", "1")
   gVarIniFile.VerDTE = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerExento", "1")
   gVarIniFile.VerExento = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerSucursal", "1")
   gVarIniFile.VerSucursal = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerNumDocHasta", "1")
   gVarIniFile.VerNumDocHasta = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerNumInterno", "1")
   gVarIniFile.VerNumInterno = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerOtrosImp", "1")
   gVarIniFile.VerOtrosImp = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "RepetirGlosa", "0")
   gVarIniFile.RepetirGlosa = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerCredArt33", "1")
   gVarIniFile.VerCredArt33 = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerFVenta", "1")
   gVarIniFile.VerFVenta = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerFUtiliz", "1")
   gVarIniFile.VerFUtiliz = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerTipoDep", "1")
   gVarIniFile.VerTipoDep = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerTipoDepHist", "1")
   gVarIniFile.VerTipoDepHist = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerValCompraHist", "1")
   gVarIniFile.VerValCompraHist = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerPatenteRol", "1")
   gVarIniFile.VerPatenteRol = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerNombreProy", "1")
   gVarIniFile.VerNombreProy = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerFechaProy", "1")
   gVarIniFile.VerFechaProy = Val(Buf)
   
  
   Buf = GetIniString(gIniFile, "Opciones", "VerMaqReg", "1")
   gVarIniFile.VerMaqReg = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerCantBoletas", "1")
   gVarIniFile.VerCantBoletas = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerDetOtrosImp", "1")
   gVarIniFile.VerDetOtrosImp = Val(Buf)

   Buf = GetIniString(gIniFile, "Opciones", "SelEmprPorRUT", "0")
   gVarIniFile.SelEmprPorRUT = Val(Buf)

   Buf = GetIniString(gIniFile, "Opciones", "VerPropIVA", "1")
   gVarIniFile.VerPropIVA = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerFechaCompra", "1")
   gVarIniFile.VerFechaCompra = Val(Buf)
      
   Buf = GetIniString(gIniFile, "Opciones", "VerValorInicial", "1")
   gVarIniFile.VerValorInicial = Val(Buf)
      
   Buf = GetIniString(gIniFile, "Opciones", "VerPjeAmortizacion", "1")
   gVarIniFile.VerPjeAmortizacion = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerFactor", "1")
   gVarIniFile.VerFactor = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerValorRazonable", "1")
   gVarIniFile.VerValorRazonable = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerRevalorizacion", "1")
   gVarIniFile.VerRevalorizacion = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLCajaOper", "1")
   gVarIniFile.VerLCajaOper = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLCajaDTE", "1")
   gVarIniFile.VerLCajaDTE = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLCajaNombre", "1")
   gVarIniFile.VerLCajaNombre = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLCajaIVAIrrec", "1")
   gVarIniFile.VerLCajaIVAIrrec = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLCajaOtrosImp", "1")
   gVarIniFile.VerLCajaOtrosImp = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerOtrosIngEgr14TER", "1")
   gVarIniFile.VerOtrosIngEgr14TER = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerCodCuenta", "1")
   gVarIniFile.VerCodCuenta = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerAreaNeg", "1")
   gVarIniFile.VerAreaNeg = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerCCosto", "1")
   gVarIniFile.VerCCosto = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerGlosaComp", "1")
   gVarIniFile.VerGlosaComp = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLibRetDTE", "1")
   gVarIniFile.VerLibRetDTE = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerLibRetSucursal", "1")
   gVarIniFile.VerLibRetSucursal = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Opciones", "VerRet3Porc", "1")
   gVarIniFile.VerRet3Porc = Val(Buf)
   
   Buf = GetIniString(gIniFile, "Config", "PathFactura", "")
   gDbFacturacion = Trim(Buf)
   
   If GetIniString(gIniFile, "Config", "Path") <> W.AppPath Then
    Call SetIniString(gIniFile, "Config", "Path", W.AppPath)
   End If

End Sub
Public Sub IniHyperCont()
   
   Call ReadIni
   'Franca 14/01/04 Se trasladan a IniEmpresa, para ser llamado desde FrmSelEmpresas
   'Call ReadParam
   'Call ReadEmpresa
   'Call LinkMdbAdm
      
   gFldOrdPlan(ORDPLAN_COD) = "Cuentas.Codigo"
   gFldOrdPlan(ORDPLAN_NOM) = "Cuentas.Nombre"
   gFldOrdPlan(ORDPLAN_DESC) = "Cuentas.Descripcion"
   
   gFindPlan(ORDPLAN_COD) = "Código"
   gFindPlan(ORDPLAN_NOM) = "Nombre"
   gFindPlan(ORDPLAN_DESC) = "Descripción"
      
   gEstadoEntidad(EE_ACTIVO) = "Activo"
   gEstadoEntidad(EE_INACTIVO) = "Inactivo"
   gEstadoEntidad(EE_BLOQUEADO) = "Bloqueado"
   
   gEstadoDoc(ED_ANULADO) = "Anulado"
   gEstadoDoc(ED_APROBADO) = "Aprobado"
   gEstadoDoc(ED_CENTRALIZADO) = "Centralizado"
   gEstadoDoc(ED_PAGADO) = "Pagado"
   gEstadoDoc(ED_PENDIENTE) = "Pendiente"
   
   gEstadoLibImp(EL_IMPRESO) = "Impreso"
   gEstadoLibImp(EL_ANULADO) = "Anulado"
   
   gClasifEnt(ENT_CLIENTE) = "Cliente"
   gClasifEnt(ENT_PROVEEDOR) = "Proveedor"
   gClasifEnt(ENT_EMPLEADO) = "Empleado"
   gClasifEnt(ENT_SOCIO) = "Socio"
   gClasifEnt(ENT_DISTRIB) = "Distribuidor"
   gClasifEnt(ENT_OTRO) = "Otro"
   
   gAtribCuentas(ATRIB_CONCILIACION).Nombre = "Cuenta Banco (Conciliación Bancaria)"
   gAtribCuentas(ATRIB_CONCILIACION).NombreCorto = "Conc"
   gAtribCuentas(ATRIB_CAPITALPROPIO).Nombre = "Capital Propio"
   gAtribCuentas(ATRIB_CAPITALPROPIO).NombreCorto = "Cap"
   gAtribCuentas(ATRIB_ACTIVOFIJO).Nombre = "Activo Fijo"
   gAtribCuentas(ATRIB_ACTIVOFIJO).NombreCorto = "AcF"
   gAtribCuentas(ATRIB_RUT).Nombre = "Documento (RUT) asociado"
   gAtribCuentas(ATRIB_RUT).NombreCorto = "Doc"
   gAtribCuentas(ATRIB_CCOSTO).Nombre = "Centro de Gestión asociado"
   gAtribCuentas(ATRIB_CCOSTO).NombreCorto = "CGes"
   gAtribCuentas(ATRIB_AREANEG).Nombre = "Área de Negocio asociada"
   gAtribCuentas(ATRIB_AREANEG).NombreCorto = "ANeg"
   gAtribCuentas(ATRIB_CAJA).Nombre = "Cuenta Caja (efectivo)"
   gAtribCuentas(ATRIB_CAJA).NombreCorto = "Caja"
   gAtribCuentas(ATRIB_14TER).Nombre = "Ajuste 14 TER"
   gAtribCuentas(ATRIB_14TER).NombreCorto = "14 TER"
   gAtribCuentas(ATRIB_PERCEPCIONES).Nombre = "Percepciones"
   gAtribCuentas(ATRIB_PERCEPCIONES).NombreCorto = "Percepcion"

   gMovActivoFijo(MOVAF_COMPRA) = "Compra"
   gMovActivoFijo(MOVAF_VENTA) = "Venta"
   gMovActivoFijo(MOVAF_BAJA) = "Baja"

   gEstadoMes(EM_NOEXISTE) = "No Existe"
   gEstadoMes(EM_ABIERTO) = "Abierto"
   gEstadoMes(EM_CERRADO) = "Cerrado"
   gEstadoMes(EM_ERRONEO) = "Erróneo"

   gTipoRetencion(TR_HONORARIOS) = "Honorarios"
   gTipoRetencion(TR_DIETA) = "Dieta"
   gTipoRetencion(TR_OTRO) = "Otro"
   
   'tipo de relación de una entidad con un documento:
   gTipoRelEnt(TRE_EMISOR) = "Emisor"
   gTipoRelEnt(TRE_RECEPTOR) = "Receptor"
   gTipoRelEnt(TRE_OTRO) = "Otro"
   
   'Franquicia Tributaria Entidad
   gFranqTribEnt(FTE_14A) = "Art. 14 A Régimen Semi Integrado"
   gFranqTribEnt(FTE_14DN3) = "Art. 14 D N° 3 Régimen Pro Pyme General"
   gFranqTribEnt(FTE_14DN8) = "Ärt. 14 D N° 8 Régimen Pro Pyme Transparente"
   gFranqTribEnt(FTE_14BN1) = "Art. 14 B N° 1 Renta Efectiva sin Balance"
   gFranqTribEnt(FTE_RRENTASPRES) = "Rentas Presuntas"
   gFranqTribEnt(FTE_OTRO) = "Otro"
  
   'Tipos de Informes IFRS
   gInformeIFRS(IFRS_ESTFIN) = "Estado de Situación Financiera Clasificado"
   gInformeIFRS(IFRS_ESTRES) = "Estado de Resultados por Función"
   gInformeIFRS(IFRS_BALEJEC) = "Estado de Situación Financiera Ejecutivo"
   gInformeIFRS(IFRS_BAL8COL) = "Balance General 8 Columnas Formato IFRS"
   
   'Tipos de ajuste en comprobantes
   gTipoAjuste(TAJUSTE_FINANCIERO) = "Financiero"
   gTipoAjuste(TAJUSTE_TRIBUTARIO) = "Tributario"
   gTipoAjuste(TAJUSTE_AMBOS) = "Ambos (Fin. y Trib.)"
   
   'Tipo de Depreciación
   gTipoDepStr(DEP_NORMAL) = "Normal"
   gTipoDepStr(DEP_ACELERADA) = "Acelerada"
   gTipoDepStr(DEP_INSTANTANEA) = "Instantánea"
   gTipoDepStr(DEP_DECIMAPARTE) = "Décima Parte"
   gTipoDepStr(DEP_DECIMAPARTE2) = "Décima Parte MT"
   
   gTipoDepLey21210Str(DEP_LEY21210_INST) = "Inst.e Inmed."
   gTipoDepLey21210Str(DEP_LEY21210_ARAUCANIA) = "Araucanía"
   
   gTipoDepLey21256Str = "Ley 21.256"
   
   gFechaInicioDepInstantanea = DateSerial(2014, 10, 1)
   gFechaTerminoDepInstantanea = DateSerial(2019, 12, 31)
   
   gFechaInicioDepDecimaParte2 = DateSerial(2020, 1, 1)
  
'   gFechaInicioDepLey21210 = DateSerial(2019, 10, 1)
'   gFechaTerminoDepLey21210 = DateSerial(2021, 12, 31)
'  Nueva fecha Victor Morales 12 nov 2020
   gFechaInicioDepLey21210 = DateSerial(2019, 10, 1)
   gFechaTerminoDepLey21210 = DateSerial(2020, 5, 31)
   
   gFechaInicioDepLey21256 = DateSerial(2020, 6, 1)
   gFechaTerminoDepLey21256 = DateSerial(2022, 12, 31)
   
   gFechaInicioSupermercados = DateSerial(2015, 1, 1)
   gFechaInicioTraspasoSupermercados = DateSerial(2015, 7, 1)
   
   gTipoSocio(0) = ""
   gTipoSocio(1) = "Persona Natural Nacional"
   gTipoSocio(2) = "Persona Natural Extranjera"
   gTipoSocio(3) = "Empresario Individual"
   gTipoSocio(4) = "Sociedad de Personas"
   gTipoSocio(5) = "Sociedad Anónima"
   gTipoSocio(6) = "Sociedad en Comandita por Acc."
   gTipoSocio(7) = "Agencia Extranj. Art. 58 N°1 L.I.R."
   gTipoSocio(8) = "Sociedad de Hecho"
   gTipoSocio(9) = "Comunidad"
   gTipoSocio(10) = "Sociedad por acciones"
   gTipoSocio(11) = "Otros"
   
   'Detalle Capital Propio Simplificado

   gTipoDetCapPropioSimpl(CPS_PARTICIPACIONES) = "Participaciones"
   gTipoDetCapPropioSimpl(CPS_DISMINUCIONES) = "Disminuciones Formales de Capital"
   gTipoDetCapPropioSimpl(CPS_GASTOSRECHAZADOS) = "Gastos Rechazados No Afectos Art. 21"
   gTipoDetCapPropioSimpl(CPS_BASEIMPONIBLE) = "Base  Imponible   Primera Categoría"
   gTipoDetCapPropioSimpl(CPS_RETDIV) = "Retiros o dividendos efectuados a propietarios"
   gTipoDetCapPropioSimpl(CPS_AUMENTOSCAP) = "Aumentos de Capital"
   gTipoDetCapPropioSimpl(CPS_GASTOSRECHNOPAGAN40) = "Gastos Rechazados Inc.1 Art. 21 No Pagan 40%"
   gTipoDetCapPropioSimpl(CPS_INRPROPIOS) = "INR Propios"
   gTipoDetCapPropioSimpl(CPS_OTROSAJUSTAUMENTOS) = "Otros Ajustes (Aumentos)"
   gTipoDetCapPropioSimpl(CPS_OTROSAJUSTDISMIN) = "Otros Ajustes (Disminuciones)"
   gTipoDetCapPropioSimpl(CPS_CAPPROPIOTRIBANOANT) = "Cap. Propio Trib. Año Anterior"
   gTipoDetCapPropioSimpl(CPS_REPPERDIDAARRASTRE) = "Reposición Pérdida de Arrastre"
   gTipoDetCapPropioSimpl(CPS_INRPROPIOSPERDIDAS) = "Pérdida por rentas exentas e ingresos no renta"
   gTipoDetCapPropioSimpl(CPS_UTILIDADESPERDIDA) = "Utilidades percibidas imputadas a la pérdida"
   gTipoDetCapPropioSimpl(CPS_INGRESODIFERIDO) = "Ingreso diferido incrementado imputado en el año"
   gTipoDetCapPropioSimpl(CPS_CTDIMPUTABLEIPE) = "CTD imputable contra Impuestos Finales (IPE)"
   gTipoDetCapPropioSimpl(CPS_INCENTIVOAHORRO) = "Incentivo al ahorro según art.14 Letra E) de la LIR"
   gTipoDetCapPropioSimpl(CPS_IDPCVOLUNTARIO) = "Base IDPC Voluntario, según art. 14 Letra A n°6 LIR"
   gTipoDetCapPropioSimpl(CPS_CREDACTFIJOS) = "Crédito por activos fijos adquiridos (art. 33 bis LIR)"
   gTipoDetCapPropioSimpl(CPS_CREDPARTICIPACIONES) = "Crédito por participaciones recibidas"
  

   gMsgLey21210 = "Estimado Usuario debido a que se publicó la Ley 21.210 Moderniza Legislación Tributaria D.O. 24.02.2020, este régimen tributario sufrirá modificaciones que saldrán en próximas versiones del sistema"

'   Call IniTraspasos         'Ya no existe FUT
   
   Call IniTraspasos20
   
   Call IniTipoDatosRemu
   
   Call IniFldExpSII
   
   Call InitBaseImponible14Ter
   
   Call InitBaseImponible14D
   
   Call InitAsistImpPrimCat
      
   gSelTipoDoc.Default = True
   
End Sub
Private Sub ReadParam()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, j As Integer
   Dim NLib As Integer
   Dim CodTV As String
   
   Call ReadComun
   
   Call AddDebug("ReadParam: pasamos ReadComun", 1)

   'Clasificación de cuentas
   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='CLASCTA' ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   Do While Rs.EOF = False
      gClasCta(vFld(Rs("Codigo"))) = vFld(Rs("Valor"))
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   'clasificación de libros: código
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM Param WHERE Tipo='TIPOLIBCOD' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTipoLibCod(i)
      gTipoLibCod(i) = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   'clasificación de libros: nombre
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM Param WHERE Tipo='TIPOLIB' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTipoLib(i)
      gTipoLib(i) = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   'clasificación de libros: id, nombre
   Set Rs = OpenRs(DbMain, "SELECT Codigo, Valor FROM Param WHERE Tipo='TIPOLIB' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTipoLibNew(i)
      gTipoLibNew(i).id = vFld(Rs("Codigo"))
      gTipoLibNew(i).Nombre = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
      Set Rs = OpenRs(DbMain, "SELECT Codigo, Valor FROM Param WHERE Tipo='TRATAMIENTO' ORDER BY Codigo")
   
   i = 1
   Do While Rs.EOF = False
   
      ReDim Preserve gTratamiento(i)
      gTratamiento(i).id = vFld(Rs("Codigo"))
      gTratamiento(i).Nombre = vFld(Rs("Valor"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   Call AddDebug("ReadParam: pasamos lectura parámetros", 1)
   
   Call ReadTipoRazFin
   
   Call AddDebug("ReadParam: pasamos lectura ReadTipoRazFin", 1)
   
   Call ReadTipoDocs
   
   Call AddDebug("ReadParam: pasamos lectura ReadTipoDocs", 1)
   
   NLib = UBound(gTipoLib)
      
   Call ReadTipoValor
   
   gLibroOficial(LIBOF_COMPRAS) = "Libro de Compras"
   gLibroOficial(LIBOF_VENTAS) = "Libro de Ventas"
   gLibroOficial(LIBOF_RETEN) = "Libro de Retenciones"
   gLibroOficial(LIBOF_DIARIO) = "Libro Diario"
   gLibroOficial(LIBOF_MAYOR) = "Libro Mayor"
   gLibroOficial(LIBOF_INVBAL) = "Libro de Inventario y Balance"
   gLibroOficial(LIBOF_TRIBUTARIO) = "Balance Tributario"
   gLibroOficial(LIBOF_CLASIFICADO) = "Balance Clasificado"
   gLibroOficial(LIBOF_COMPYSALDOS) = "Balance de Comprobación y Saldos"
   gLibroOficial(LIBOF_ESTRESCLASIF) = "Estado de Resultado Clasificado"
   gLibroOficial(LIBOF_ESTRESCOMP) = "Estado de Resultado Comparativo"
   gLibroOficial(LIBOF_ESTRESMENSUAL) = "Estado de Resultado Mensual"
   gLibroOficial(LIBOF_INGEGR) = "Libro de Ingresos y Egresos"
   
   gTipoOperCaja(0) = ""
   gTipoOperCaja(TOPERCAJA_INGRESO) = "Ingreso"
   gTipoOperCaja(TOPERCAJA_EGRESO) = "Egreso"
      
      
   gTipoDocCajaOtros(0) = ""
   gTipoDocCajaOtros(LIBCAJA_OTROSING) = "Otros Ingresos"
   gTipoDocCajaOtros(LIBCAJA_OTROSEGR) = "Otros Egresos"
   
      
   Call AddDebug("ReadParam: pasamos lectura var globales", 1)
      
   Call ReadIndices
   
   Call AddDebug("ReadParam: Nos vamos", 1)
      
End Sub
Public Function ReadTipoDocs()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
 
   'tipos de docs, independiente de los libros
   
   '2814014 pipe
   ReDim gTipoDoc(11) '2814014
   'ReDim gTipoDoc(10)
   'fin 2814014
   
   Q1 = "SELECT Id, TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TieneAfecto, TieneExento, IngresarTotal, TieneNumDocHasta, TieneCantBoletas, ExigeRUT, "
   Q1 = Q1 & " EsRebaja, DocBoletas, DocImpExp, CodDocSII, CodDocDTESII, AceptaPropIVA "
   Q1 = Q1 & " FROM TipoDocs WHERE Atributo='ACTIVO' ORDER BY TipoLib, TipoDoc"
   Set Rs = OpenRs(DbMain, Q1)
      
   i = 0
   Do While Rs.EOF = False
   
      If i > UBound(gTipoDoc) Then
         ReDim Preserve gTipoDoc(i + 10)
      End If
      
      gTipoDoc(i).id = vFld(Rs("Id"))
      gTipoDoc(i).TipoLib = vFld(Rs("TipoLib"))
      gTipoDoc(i).TipoDoc = vFld(Rs("TipoDoc"))
      gTipoDoc(i).Nombre = vFld(Rs("Nombre"))
      gTipoDoc(i).Diminutivo = vFld(Rs("Diminutivo"))
      gTipoDoc(i).Atributo = vFld(Rs("Atributo"))
      gTipoDoc(i).TieneAfecto = vFld(Rs("TieneAfecto"))
      gTipoDoc(i).TieneExento = vFld(Rs("TieneExento"))
      gTipoDoc(i).IngresarTotal = vFld(Rs("IngresarTotal"))
      gTipoDoc(i).TieneNumDocHasta = vFld(Rs("TieneNumDocHasta"))
      gTipoDoc(i).TieneCantBoletas = vFld(Rs("TieneCantBoletas"))
      gTipoDoc(i).ExigeRUT = vFld(Rs("ExigeRUT"))
      gTipoDoc(i).EsRebaja = vFld(Rs("EsRebaja"))
      gTipoDoc(i).DocBoletas = vFld(Rs("DocBoletas"))
      gTipoDoc(i).TipoDocLAU = -1
      gTipoDoc(i).DocImpExp = vFld(Rs("DocImpExp"))
      gTipoDoc(i).CodDocSII = vFld(Rs("CodDocSII"))
      gTipoDoc(i).CodDocDTESII = vFld(Rs("CodDocDTESII"))
      gTipoDoc(i).AceptaPropIVA = vFld(Rs("AceptaPropIVA"))
      
      
#If DATACON = 1 Then
      'asignamos TipoDoc LAU
      If gTipoDoc(i).TipoLib = LIB_COMPRAS Then
      
         Select Case gTipoDoc(i).Diminutivo
         
            Case "FAC"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_FACT
      
            Case "NDC"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_NOTADEB
      
            Case "NCC"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_NOTACRED
      
            Case "FCC"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_FACTCOMP
      
            Case "OTC"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_OTRO
      
            Case "FCE"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_FACTEXEN
               
            Case "IMP"
               gTipoDoc(i).TipoDocLAU = LAU_COMP_FACTIMP
               
            Case Else
               gTipoDoc(i).TipoDocLAU = -1
               
         End Select
      
      ElseIf gTipoDoc(i).TipoLib = LIB_VENTAS Then
      
         Select Case gTipoDoc(i).Diminutivo
         
            Case "FAV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_FACT
      
            Case "NDV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_NOTADEB
      
            Case "NCV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_NOTACRED
      
            Case "FVE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_FACTEXEN
      
            Case "OTV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_OTRO
      
            Case "FCV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_FACTCOMP
               
            Case "LFV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_LIQFACT
               
           Case "BOV"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_BOLETA
               
            Case "DVB"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_DEVBOLETA
               
            Case "EXP"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_FACTEXP
               
            Case "NCE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_NCREDEXP
            
            Case "NDE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_NDEBEXP
            
            Case "BOE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_BOLEXENTA
            
            Case "VEM"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_VTAMENOR
               
            Case "VPE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_VALEPAGOELECTR
               
'              '2814014
            Case "VPEE"
               gTipoDoc(i).TipoDocLAU = LAU_VENTA_BOLVENTAEXENTA
'               'fin 2814014
            Case Else
               gTipoDoc(i).TipoDocLAU = -1
               
         End Select
         
      End If
#End If
      
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)


End Function


'Genera un nuevo codigo de cuenta y además entrega su último hermano
Public Function NewCodigoCta(PCta As Cuenta_t, CodHermano As String, Optional ByVal MaxMenos As Integer = 0) As String
   Dim Codigo As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim LastCod As Integer
   Dim NewCodigo As String
   Dim NewCta As Integer
   Dim i As Integer
   
   Codigo = Left(PCta.Codigo, gNiveles.Inicio(PCta.Nivel) - 1) & "9" & String(gNiveles.Largo(PCta.Nivel) - 1, "8")
   Codigo = Left(Codigo & String(3 * 5, "0"), Len(PCta.Codigo))
   
   Q1 = "SELECT Max(Codigo) as CodMax FROM Cuentas WHERE Nivel=" & PCta.Nivel & " AND idPadre=" & PCta.IdPadre & " AND Codigo < '" & Codigo & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   CodHermano = vFld(Rs("CodMax"))
   Call CloseRs(Rs)
   
   'NewCodigo = Mid(CodHermano, gNiveles.Largo(PCta.NivelPadre) * PCta.NivelPadre + 1, gNiveles.Largo(PCta.Nivel))
   NewCodigo = Mid(CodHermano, gNiveles.Inicio(PCta.Nivel), gNiveles.Largo(PCta.Nivel))
   
   NewCta = Val(NewCodigo) + 1
      
   NewCodigo = Right(String(gNiveles.Largo(PCta.Nivel), "0") & NewCta, gNiveles.Largo(PCta.Nivel))
   
   If PCta.Codigo = "*" Then 'Primer nivel
      Codigo = NewCodigo & "-" & Mid(gFmtCodigoCta, InStr(gFmtCodigoCta, "-") + 1)
   Else
      'Codigo = Left(PCta.Codigo, gNiveles.Largo(PCta.NivelPadre) * PCta.NivelFather) & NewCodigo '& Mid(PCta.Codigo, gNiveles.Largo(PCta.NivelFather) * PCta.Nivel + 1)
      Codigo = Left(PCta.Codigo, gNiveles.Inicio(PCta.Nivel) - 1) & NewCodigo '& Mid(PCta.Codigo, gNiveles.Largo(PCta.NivelFather) * PCta.Nivel + 1)
      If PCta.Nivel <> gLastNivel Then
         'Codigo = Codigo & Mid(PCta.Codigo, gNiveles.Largo(PCta.NivelFather) * PCta.Nivel + 1)
         Codigo = Codigo & Mid(PCta.Codigo, gNiveles.Inicio(PCta.Nivel + 1))
      End If
      Codigo = Format(Codigo, gFmtCodigoCta)
      
   End If
  
   NewCodigoCta = Codigo
   
   Q1 = "SELECT IdCuenta FROM Cuentas WHERE Codigo = '" & VFmtCodigoCta(Codigo) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'ya existe
      MsgBox1 "No es posible definir más sub-cuentas bajo esta cuenta.", vbExclamation + vbOKOnly
      NewCodigoCta = ""
   End If
   Call CloseRs(Rs)
   
   Exit Function
   
   For i = 1 To 10
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Codigo = '" & VFmtCodigoCta(Codigo) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then 'ya existe
         Call CloseRs(Rs)
         Codigo = NewCodigoCta(PCta, CodHermano, i)
      Else
         NewCodigoCta = Codigo
         Exit Function
      End If
      Call CloseRs(Rs)
   Next i
   
   MsgBox1 "No es posible definir más sub-cuentas bajo esta cuenta.", vbExclamation + vbOKOnly
   NewCodigoCta = ""
   
End Function
Public Function VFmtCodigoCta(ByVal CodCta As String) As String
   CodCta = Trim(Replace(CodCta, "-", ""))
   VFmtCodigoCta = CodCta
   
End Function

Public Function FmtCodCuenta(ByVal CodCuenta) As String
   FmtCodCuenta = Format(CodCuenta, gFmtCodigoCta)
   
End Function
Public Function FmtCodIFRS(ByVal CodIFRS) As String
   FmtCodIFRS = Format(CodIFRS, gFmtCodigoIFRS)
   
End Function
Public Function VFmtCodigoIFRS(ByVal CodIFRS As String) As String
   CodIFRS = Trim(Replace(CodIFRS, "-", ""))
   VFmtCodigoIFRS = CodIFRS
   
End Function
Sub KeyCodCta(KeyAscii As Integer)
   
   If Not (IsNumeric(Chr(KeyAscii))) And Chr(KeyAscii) <> "-" And KeySys(KeyAscii) = False Then
      Beep
      KeyAscii = 0
   End If

End Sub
Public Function IniEmpresa() As Boolean
   Dim IdCompAperTrib As Long
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim DbName As String
   
   IniEmpresa = False
   'EsEmpresaNueva = False
   
#If DATACON = 1 Then       'Access
   
   'abrimos la base de datos de la empresa
   Call AddDebug("IniEmpresa: a OpenDbEmp", 2)
   If gEmprSeparadas Then
      If OpenDbEmp() = False Then
       '  End
         Exit Function
      End If
   End If
   
   Call AddDebug("IniEmpresa: a ChkDbInfo", 2)
   If ChkDbInfo(DbMain, gEmpresa.Rut, gEmpresa.Ano, gEmpresa.id) = False Then
      Call CloseDb(DbMain)
      Exit Function
   End If
   
   'linkeamos las tablas por si se movieron
   If gEmprSeparadas Then
      Call AddDebug("IniEmpresa: a LinkMdbAdm", 2)
      Call LinkMdbAdm
   End If
   
   
   'modificaciones a la base de datos de acuerdo a la versión (debe estar después de LinkMdbAdm)
   Call AddDebug("IniEmpresa: a CorrigeBase", 2)
   Call CorrigeBase
   
   '********* Agrega el campo tratamiento si no lo tiene **********
   On Error Resume Next
    ERR.Clear
    DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
    If ExistFile(DbName) Then
        Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
        Set Tbl = DbMain.TableDefs("Documento")
        
        'agregamos campo Tratamiento a tabla Documento
        ERR.Clear
        Tbl.Fields.Append Tbl.CreateField("Tratamiento", dbLong)
        If ERR = 0 Then
          Tbl.Fields.Refresh
        End If
    End If
   
    ERR.Clear

    Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
    Set Tbl = DbMain.TableDefs("Documento")
    
    'agregamos campo Tratamiento a tabla Documento
    ERR.Clear
    Tbl.Fields.Append Tbl.CreateField("Tratamiento", dbLong)
    If ERR = 0 Then
      Tbl.Fields.Refresh
    End If

    '************ FINNNNNN ********************
   
#Else
   'modificaciones a la base de datos de acuerdo a la versión
   Call CorrigeBaseSQLServer
   
   '627184
   Call CampoCodSIIDTE
   '627184
#End If
      
   'estas dos líneas estaban en CrearNuevoAno
   
   Call GenCompAperSinMovs(1, gEmpresa.id, gEmpresa.Ano, IdCompAperTrib)
   
   Call InsertParamEmpBas(gEmpresa.id, gEmpresa.Ano)
   
   '3410269
   Call ModConBoletaVPEE
   '3410269
   
   'inicializamos datos básicos del sistema
   Call AddDebug("IniEmpresa: a ReadParam", 2)
   Call ReadParam
   
   gAtribCuentas(ATRIB_14TER).Nombre = IIf(gEmpresa.Ano < 2020, "Ajuste 14 TER", "Ajuste 14D N°3 y 8 LIR")
   gAtribCuentas(ATRIB_14TER).NombreCorto = IIf(gEmpresa.Ano < 2020, "14 TER", "14 D")

   
   'inicializamos datos básicos de la empresa
   Call AddDebug("IniEmpresa: a ReadEmpresa", 2)
   Call ReadEmpresa
   
   If gFunciones.ProporcionalidadIVA Then
      Call InitPropIVA
      Call PropIVA_UpdateTblTotMensual
   End If
   
   Call InitCtasAjustesExtraCont
   Call InitCtasAjustesExtraContRLI
   
   '2850275
    'Call EmpresasLpRemu
   'fin 2850275
   
   '633824 ffv
'   Call CampoDocOtroEsCargoRem
   '633824 ffv
   
   
   '14520904
  If gEmpresa.NuevoAno = False Then
   
     Call CorregirSaldosAnoAnterior
        '14520904

     Call CorrigeDocPendientesAñoAnterior(False)
    End If
  
   IniEmpresa = True
   
   Call AddDebug("IniEmpresa: nos vamos OK", 1)

End Function
Public Sub ReadEmpresa()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Ano As String
   Dim i As Integer
   Dim MaxFecha As Long
   Dim MesActual As Integer
   Dim IdCompAper As Long
   Dim IdCompAperTrib As Long
   Dim TotalDebeAper As Double
   Dim TotalDebeAperTrib As Double
     
     
   Call ReadDatosBasEmpresa
  
   'veamos si ya se generó comprobante de apertura
   Q1 = "SELECT IdCompAper, IdCompAperTrib "
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      IdCompAper = vFld(Rs("IdCompAper"))
      IdCompAperTrib = vFld(Rs("IdCompAperTrib"))
   End If
   
   Call CloseRs(Rs)
   
   If IdCompAper <> 0 Then
      'vemos si este comprobante existe realmente
      Q1 = "SELECT Tipo, TotalDebe FROM Comprobante WHERE IdComp = " & IdCompAper
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then    'no existe
         IdCompAper = 0
         
         Q1 = "UPDATE EmpresasAno SET IdCompAper = 0 "
         Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      ElseIf vFld(Rs("Tipo")) <> TC_APERTURA Then   'existe el comprobante pero no es de apertura
      
         IdCompAper = 0
         
         Q1 = "UPDATE EmpresasAno SET IdCompAper = 0 "
         Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      Else
         TotalDebeAper = vFld(Rs("TotalDebe"))
            
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If IdCompAperTrib <> 0 Then
      'vemos si este comprobante existe realmente
      Q1 = "SELECT Tipo, TotalDebe FROM Comprobante WHERE IdComp = " & IdCompAperTrib
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then    'no existe
         IdCompAperTrib = 0
         
         Q1 = "UPDATE EmpresasAno SET IdCompAperTrib = 0 "
         Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      ElseIf vFld(Rs("Tipo")) <> TC_APERTURA Then   'existe el comprobante pero no es de apertura
      
         IdCompAperTrib = 0
         
         Q1 = "UPDATE EmpresasAno SET IdCompAperTrib = 0 "
         Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
            
       Else
         TotalDebeAperTrib = vFld(Rs("TotalDebe"))
           
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   'si tiene año anterior y no se ha generado comprobante de apertura, lo generamos
   'también traemos los docs pendientes y activos fijos con residual del año anterior, si hay
   If (gEmpresa.TieneAnoAnt Or gEmpresa.TieneAnoAntAccess) And gEmpresa.DebeGenCompAp Then
'      If IdCompAper = 0 Or IdCompAperTrib = 0 Then
      
         If GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano) = True Then
            'MsgBox1 "Se generó el Comprobante de Apertura.", vbInformation + vbOKOnly
            
            If Not gEmpresa.TieneAnoAntAccess Then   'docs pendientes y act. fijo residual ya fueron generados en CrearNuevoAnoSQLFromAccess
               Call GenDocsPendientes(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, True)
               'Call GenDocsFullPendientes(gEmpresa.Id, gEmpresa.Rut, gEmpresa.Ano, True, True)
               Call GenActFijoResidual(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, True)
            End If
            
            gEmpresa.DebeGenCompAp = False
            Q1 = "UPDATE ParamEmpresa SET Codigo = 0 "
            Q1 = Q1 & " WHERE Tipo = 'INITAÑO'"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
         End If
            
   End If
   
   'leemos saldo Libro de Caja año anterior
   gSaldoLibroCajaAnoAnt = 0
   If gEmpresa.TieneAnoAnt Then
      Q1 = "SELECT SaldoLibroCaja "
      Q1 = Q1 & " FROM EmpresasAno "
      Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano - 1
   
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         gSaldoLibroCajaAnoAnt = vFld(Rs("SaldoLibroCaja"))
      End If
      
      Call CloseRs(Rs)
   End If
   
   Call ReadCtasAjustesExtraCont
   Call ReadCtasAjustesExtraContRLI
      
End Sub
Public Sub ReadDatosBasEmpresa()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Ano As String
   Dim i As Integer
   Dim MaxFecha As Long
   Dim MesActual As Integer
   Dim InitAno As String
   Dim IdCompAper As Long
   Dim ValUTM As Double
   Dim CodInitAno As Integer
      
   Call AddDebug("ReadDatosBasEmpresa: llegamos", 1)
   
      
   gEmpresa.Direccion = ""
   gEmpresa.Telefono = ""
   gEmpresa.RazonSocial = ""
   gEmpresa.Comuna = ""
   gEmpresa.Ciudad = ""
   gEmpresa.Giro = ""
   gEmpresa.CodActEcono = ""
   gEmpresa.RepConjunta = False
   gEmpresa.RutRepLegal1 = ""
   gEmpresa.RepLegal1 = ""
   gEmpresa.RutRepLegal2 = ""
   gEmpresa.RepLegal2 = ""
   gEmpresa.Opciones = 0
   gEmpresa.Franq14Ter = 0
   gEmpresa.RentaAtribuida = 0
   gEmpresa.SemiIntegrado = 0
   gEmpresa.R14ASemiIntegrado = 0
   gEmpresa.SocProfSegCat = 0
   gEmpresa.ObligaLibComprasVentas = 0
   
   '2861570
   gEmpresa.RutContador = ""
   'fin 2861570
   
   '2913643
   gEmpresa.Villa = ""
   gEmpresa.Celular = "0"
   gEmpresa.CodArea = "0"
   '2913643

   'leemos los datos de la empresa en el único registro de esta tabla
   Q1 = "SELECT NombreCorto, Calle, Numero, Dpto, Telefonos, RazonSocial, ApMaterno, Nombre "
   Q1 = Q1 & ", Giro, CodActEconom, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2"
   Q1 = Q1 & ", Region, Regiones.Comuna, Ciudad, Opciones, TipoContrib "
   Q1 = Q1 & ", Franq14ter, FranqRentaAtribuida, FranqSemiIntegrado, FranqProPymeGeneral, FranqProPymeTransp, FranqSocProfSegCat, ObligaLibComprasVentas, Franq14ASemiIntegrado "
   
   '2861570
   Q1 = Q1 & ",RutContador "
   'fin2861570
   
   '2913643
   Q1 = Q1 & ",villa,celular,codArea "
   '2913643
   
   Q1 = Q1 & "  FROM Empresa LEFT JOIN Regiones ON Empresa.Comuna=Regiones.id"
   Q1 = Q1 & "  WHERE Empresa.Id = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then
      'ES LA PRIMERA Y SE HACE INSERT EN ESTA TABLA
      Q1 = "INSERT INTO Empresa (id, Ano, Rut, NombreCorto, RazonSocial) VALUES ("
      Q1 = Q1 & gEmpresa.id
      Q1 = Q1 & "," & gEmpresa.Ano
      Q1 = Q1 & ",'" & gEmpresa.Rut & "'"
      Q1 = Q1 & ",'" & gEmpresa.NombreCorto & "'"
      Q1 = Q1 & ",'" & gEmpresa.NombreCorto & "')"
      Call ExecSQL(DbMain, Q1)
      
   Else
      gEmpresa.Direccion = vFld(Rs("Calle"), True) & " " & vFld(Rs("Numero"), True) & " " & vFld(Rs("Dpto"), True)
      gEmpresa.Telefono = vFld(Rs("Telefonos"), True)
      gEmpresa.RazonSocial = vFld(Rs("RazonSocial"), True) & " " & vFld(Rs("ApMaterno"), True) & " " & vFld(Rs("Nombre"), True)
      gEmpresa.Region = vFld(Rs("Region"))
      gEmpresa.Comuna = FCase(vFld(Rs("Comuna"), True))
      gEmpresa.Ciudad = vFld(Rs("Ciudad"), True)
      gEmpresa.Giro = vFld(Rs("Giro"), True)
      gEmpresa.CodActEcono = vFld(Rs("CodActEconom"), True)
      gEmpresa.RepConjunta = vFld(Rs("RepConjunta"))
      gEmpresa.RutRepLegal1 = vFld(Rs("RutRepLegal1"))
      gEmpresa.RepLegal1 = vFld(Rs("RepLegal1"), True)
      gEmpresa.RutRepLegal2 = vFld(Rs("RutRepLegal2"))
      gEmpresa.RepLegal2 = vFld(Rs("RepLegal2"), True)
      gEmpresa.Opciones = vFld(Rs("Opciones"))
      gEmpresa.Franq14Ter = vFld(Rs("Franq14ter"))
      gEmpresa.RentaAtribuida = vFld(Rs("FranqRentaAtribuida"))
      gEmpresa.SemiIntegrado = vFld(Rs("FranqSemiIntegrado"))
      gEmpresa.ProPymeGeneral = vFld(Rs("FranqProPymeGeneral"))
      gEmpresa.ProPymeTransp = vFld(Rs("FranqProPymeTransp"))
      gEmpresa.SocProfSegCat = vFld(Rs("FranqSocProfSegCat"))
      gEmpresa.R14ASemiIntegrado = vFld(Rs("Franq14ASemiIntegrado"))
      gEmpresa.ObligaLibComprasVentas = vFld(Rs("ObligaLibComprasVentas"))
      gEmpresa.TipoContrib = vFld(Rs("TipoContrib"))
      FProPymeGeneral = IIf(vFld(Rs("FranqProPymeGeneral")), True, False)
      FProPymeTransp = IIf(vFld(Rs("FranqProPymeTransp")), True, False)
      
        '2861570
        gEmpresa.RutContador = vFld(Rs("RutContador"))
        'fin 2861570
        
        '2913643
        gEmpresa.Villa = vFld(Rs("Villa"))
        gEmpresa.Celular = vFld(Rs("Celular"))
        gEmpresa.CodArea = vFld(Rs("CodArea"))
        '2913643
      
   End If
   
   Call CloseRs(Rs)
   
   Call AddDebug("ReadDatosBasEmpresa: pasamos datos empresa", 1)
   
   
   'Niveles Plan de Cuentas empresa
   Call ConfigNiveles
   
   Call AddDebug("ReadDatosBasEmpresa: pasamos ConfigNiveles", 1)
   
   'Tipo y periodo Correlativo Comprobante
   gTipoCorrComp = TCC_TIPOCOMP    'default
   gPerCorrComp = TCC_MENSUAL       'default
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='TCORRCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gTipoCorrComp = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='PCORRCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gPerCorrComp = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   gEstadoNewComp = EC_PENDIENTE
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='ESTADOCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gEstadoNewComp = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   gAbrirMesesParalelo = False
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='MESPARALEL'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gAbrirMesesParalelo = (vFld(Rs("Valor")) <> 0)
   End If

   Call CloseRs(Rs)
   
   gCompAnuladoLibDiario = False
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='VERCOMPANU'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCompAnuladoLibDiario = (vFld(Rs("Valor")) <> 0)
   End If

   Call CloseRs(Rs)
   
   gPrtMovDetOpt = 0
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='PRTMOVDET'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gPrtMovDetOpt = vFld(Rs("Valor"))
   End If

   Call CloseRs(Rs)
   
   gImpResCent = False
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='IMPRESCENT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gImpResCent = (vFld(Rs("Valor")) <> 0)
   End If

   Call CloseRs(Rs)
   
   gDtCompCent = 0
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='DTCOMPCENT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gDtCompCent = vFld(Rs("Valor"))
   End If

   Call CloseRs(Rs)

   gDayDtCompCent = 0
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='DYCOMPCENT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gDayDtCompCent = vFld(Rs("Valor"))
   End If

   Call CloseRs(Rs)
   
   gAFMesCompleto = False
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='AFMESCOMPT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gAFMesCompleto = (vFld(Rs("Valor")) <> 0)
   End If

   Call CloseRs(Rs)
   
   gTituloTipoComp = False
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='TITTIPCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gTituloTipoComp = (vFld(Rs("Valor")) <> 0)
   End If

   Call CloseRs(Rs)
     
   If gEmpresa.Ano < 2012 Then
      gMaxUTMCred33 = 650
      
   Else     '2012 en adelante
      gMaxUTMCred33 = 500
      
   End If
    
   If GetValMoneda("UTM", ValUTM, DateSerial(gEmpresa.Ano, 12, 1)) = False Then
      gMaxUTMCred33_Pesos = 0
      
   Else
      gMaxUTMCred33_Pesos = gMaxUTMCred33 * ValUTM
      
   End If
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='MAXCRED33'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then 'está
      gMaxCred33 = vFld(Rs("Valor"))
   Else
      Q1 = "INSERT INTO ParamEmpresa(Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES('MAXCRED33', 0, -1, " & gEmpresa.id & "," & gEmpresa.id & ")"    '-1 para indicar que el usuario no lo ha ingresado
      Call ExecSQL(DbMain, Q1)
      gMaxCred33 = -1
   End If

   Call CloseRs(Rs)
   
   'cuentas básicas
   
   'limpiamos las cuentas
   Call CleanCtasBas(gCtasBas)
   
   'ahora las leemos
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAIVACRED'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaIVACred = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAIVADEB'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaIVADeb = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAOIMPCRE'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaOtrosImpCred = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAOIMPDEB'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaOtrosImpDeb = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAIMPRET'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaImpRet = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTANETORET'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaNetoHon = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTANETODIE'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaNetoDieta = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
      
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAIMPUNIC'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaImpUnico = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAPATRIM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaPatrimonio = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTARESEJE'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaResEje = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)

   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAPAGOFAC'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaPagoFacturas = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTACOBFAC'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaCobFacturas = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTACREDIVA'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaCredIVA = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAIVAIRRE'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaIVAIrrec = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTARET3PRC'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaRet3Porc = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTA3CENREM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCta3PorcCentraRem = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   
   ' 2699582 PIPE
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
    Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAPPMOBLI'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaPpmObligatorio = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
    Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAPPMVOLU'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaPpmVoluntario = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   End If
   'FIN 2699582
    
   '2855046 ffv
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAODFACTI'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaOdfActivo = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CTAODFPASI'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gCtasBas.IdCtaOdfPasivo = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
    
   '2855046 ffv
   Call AddDebug("ReadDatosBasEmpresa: pasamos ParamEmpresa", 1)

   
   'veamos si la empresa tiene historia (año anterior a partir del cual se generó este año)
   Q1 = "SELECT Valor, Codigo FROM ParamEmpresa WHERE Tipo='INITAÑO'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      InitAno = vFld(Rs("Valor"))
      CodInitAno = vFld(Rs("Codigo"))
   End If
   
   Call CloseRs(Rs)
   
   gEmpresa.TieneAnoAnt = False
   
   If InitAno = "EMPHISTORIA" Then
      gEmpresa.TieneAnoAnt = True
      gEmpresa.TieneAnoAntAccess = False
      gEmpresa.DebeGenCompAp = IIf(CodInitAno <> 0, True, False)  'Indica si es la primera vez que ingresa al año, para que el sistema ofrezca generar Comp. Apertura, si este no ha sido generado y traer docs año aneterior
   ElseIf InitAno = "EMPHISTACC" Then   'estamos en SQL y el año anterior está en Access
      gEmpresa.TieneAnoAntAccess = True
      gEmpresa.DebeGenCompAp = IIf(CodInitAno <> 0, True, False)  'Indica si es la primera vez que ingresa al año, para que el sistema ofrezca generar Comp. Apertura, si este no ha sido generado y traer docs año aneterior
   Else
      'vemos si existe archivo de año anterior
      If gDbType = SQL_ACCESS Then
         If ExistFile(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb") Then
            Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano)"
            Q1 = Q1 & " VALUES( 'INITAÑO', 0, 'EMPHISTORIA'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
            Call ExecSQL(DbMain, Q1)
            gEmpresa.TieneAnoAnt = True
            gEmpresa.DebeGenCompAp = IIf(CodInitAno <> 0, True, False)  'Indica si es la primera vez que ingresa al año, para que el sistema ofrezca generar Comp. Apertura, si este no ha sido generado y traer docs año aneterior
         End If
      Else   'Vemos si hay año anterior ingresado al sistema
         Q1 = "SELECT FApertura FROM EmpresasAno WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano)"
            Q1 = Q1 & " VALUES( 'INITAÑO', 0, 'EMPHISTORIA'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
            Call ExecSQL(DbMain, Q1)
            gEmpresa.TieneAnoAnt = True
            gEmpresa.DebeGenCompAp = IIf(CodInitAno <> 0, True, False)  'Indica si es la primera vez que ingresa al año, para que el sistema ofrezca generar Comp. Apertura, si este no ha sido generado y traer docs año aneterior
         Else
            gEmpresa.TieneAnoAnt = False
         End If
         Call CloseRs(Rs)
      End If
   End If
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='VALORIVA'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gIVA = Val((vFld(Rs("Valor"))))
   End If
   Call CloseRs(Rs)
  
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='IMP1CAT'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gImpPrimCategoria = Val((vFld(Rs("Valor"))))
   End If
   Call CloseRs(Rs)
  
   gCredArt33 = 0.08
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CREDART33'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gCredArt33 = Val((vFld(Rs("Valor"))))
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('CREDART33',0,'0.08')"
      Call ExecSQL(DbMain, Q1)
   End If
   Call CloseRs(Rs)
   
   gCredArt33_2014 = 0.08
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CREDART334'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gCredArt33_2014 = Val((vFld(Rs("Valor"))))
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('CREDART334',0,'0.08')"
      Call ExecSQL(DbMain, Q1)
   End If
   Call CloseRs(Rs)
    
   gCredArt33_2015 = 0.06
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='CREDART335'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gCredArt33_2015 = Val((vFld(Rs("Valor"))))
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('CREDART335',0,'0.06')"
      Call ExecSQL(DbMain, Q1)
   End If
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='PLANCTAS'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gPlanCuentas = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
   gOcultarImpAdicDescont = 0
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='NOIMPADESC'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gOcultarImpAdicDescont = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
   'Notas
   Q1 = "SELECT Nota, Incluir, IncluirInfo FROM Notas WHERE Tipo='ART100'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gNotaArt100.TxtNota = vFld(Rs("Nota"), True)
      gNotaArt100.IncluirBal = (Abs(vFld(Rs("Incluir"))) And C_INCNOTABAL) <> 0
      gNotaArt100.IncluirLib = (Abs(vFld(Rs("Incluir"))) And C_INCNOTALIB) <> 0
      gNotaArt100.IncluirInfo = Abs(vFld(Rs("IncluirInfo"))) <> 0
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Nota, Incluir, IncluirInfo FROM Notas WHERE Tipo='NOTAESP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gNotaEspecial.TxtNota = vFld(Rs("Nota"), True)
      gNotaEspecial.IncluirBal = (Abs(vFld(Rs("Incluir"))) And C_INCNOTABAL) <> 0
      gNotaEspecial.IncluirLib = (Abs(vFld(Rs("Incluir"))) And C_INCNOTALIB) <> 0
      gNotaEspecial.IncluirInfo = Abs(vFld(Rs("IncluirInfo"))) <> 0
   End If
   
   Call CloseRs(Rs)
   

   
   Call AddDebug("ReadDatosBasEmpresa: pasamos Notas", 1)
   
   'estado meses
   
   Set Rs = OpenRs(DbMain, "SELECT Mes, Estado FROM EstadoMes WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " ORDER BY Mes desc")
   
   MesActual = 0
   i = 0
   
   'calculamos el mes actual
   Do While Rs.EOF = False
      i = i + 1
      If vFld(Rs("Estado")) = EM_ABIERTO Then
         MesActual = vFld(Rs("Mes"))
         Exit Do
      End If
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If i = 0 Then    'la tabla está vacía
      Call LlenarTablaMeses("")
      MesActual = 1          'parte con enero
      AbrirMes (MesActual)
      
   'ElseIf MesActual = 0 Then   'no hay ningún mes abierto => se terminó el año
   
   End If
   
   gColores(1) = COLOR_VERDEOSCURO
   gColores(2) = COLOR_AZULOSCURO
   gColores(3) = COLOR_MORADO
   gColores(4) = vbBlack
   gColores(5) = vbBlue
   
   'Colores
   Q1 = "SELECT Nivel, Color FROM Colores WHERE IdEmpresa = " & gEmpresa.id & " ORDER BY Nivel"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      gColores(vFld(Rs("Nivel"))) = vFld(Rs("Color"))
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   
   
   Call AddDebug("ReadDatosBasEmpresa: pasamos Colores", 1)
   
   '2860036
   'Membretes
   Q1 = "SELECT TituloMembrete1, TituloMembrete2, Texto1, Texto2 FROM Membrete "
   Q1 = Q1 & " where IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      gMembrete.TxtTitMembrete1 = vFld(Rs("TituloMembrete1"))
      gMembrete.TxtTitMembrete2 = vFld(Rs("TituloMembrete2"))
      gMembrete.TxtTexto1 = vFld(Rs("Texto1"))
      gMembrete.TxtTexto2 = vFld(Rs("Texto2"))
   End If
   
   Call CloseRs(Rs)
      
    Call AddDebug("ReadDatosBasEmpresa: pasamos Membretes", 1)
   'fin  2860036
   
   '3387590 Permite Afecto 0
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='AFECTOCERO'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      gAfectoCero = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
    Call AddDebug("ReadDatosBasEmpresa: pasamos Afecto 0", 1)
   '3387590
   
   'inicializo Foliación
   Call Foliacion
   
   Call AddDebug("ReadDatosBasEmpresa: pasamos Foliacion", 1)
   
   'llena arreglo con factores de corrección monetaria
   Call FillCorrMon(gEmpresa.Ano)
   
   Call AddDebug("ReadDatosBasEmpresa: nos vamos", 1)
   
End Sub
Public Sub LinkMdbAdm(Optional ByVal bForce As Boolean = 0)
   Dim DbComun As String
   Dim ConnStr As String
   Dim Tm As Double

#If DATACON = 1 Then       'Access
   
   Tm = CDbl(Now)
   DbComun = gDbPath & "\" & BD_COMUN
   
   'ConnStr = "PWD=" & PASSW_LEXCONT & ";"
   'ConnStr = Mid(gComunConnStr, 2)  ' sin el ; del inicio
   
   If bForce = False Then
      bForce = Val(GetIniString(gIniFile, "Config", "ReLink", "0"))
   End If

   ConnStr = gComunConnStr
   
   Call LinkMdbTable(DbMain, DbComun, "CodActiv", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Empresas", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "EmpresasAno", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Equivalencia", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Impuestos", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Monedas", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Param", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PlanAvanzado", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PlanBasico", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PlanIntermedio", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Regiones", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Timbraje", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoValor", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Usuarios", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "UsuarioEmpresa", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "Perfiles", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "IPC", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "TipoDocs", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "ControlEmpresa", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PcUsr", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "PlanCuentasSII", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "FactorActAnual", , bForce, , ConnStr, True)
   Call LinkMdbTable(DbMain, DbComun, "CapPropioSimplAnual", , bForce, , ConnStr, True)
  
   If gFunciones.IFRS Then    'ya no se usa
      Call LinkMdbTable(DbMain, DbComun, "IFRS_PlanIFRS", , bForce, , ConnStr, True)
   End If
   
   'gLinkF22 = LinkDbfTable(DbMain, gHRPath & "\PAR", "NContrib.dbf", "HR_NContrib", "FoxPro 2.0", , False)
   If Not gLinkF22 Then
      gLinkF22 = ExistFile(gHRPath & "\PAR\BD_HR_admin.mdb")
   End If
   
   gPathForm22 = "\FORM22"
   gPathPlan22 = "\PLAN22"
   
   
   If gFunciones.RazFinancieras Then
      'Call LinkMdbTable(DbMain, DbComun, "CuentasRazon", , , , ConnStr)    'ya no se linkea porque está en la DB de la empresa
      Call LinkMdbTable(DbMain, DbComun, "RazonesFin", , bForce, , ConnStr)
   End If
   
   If gFunciones.ExpFUT Then
      gLinkParFUT = LinkDbfTable(DbMain, gHRPath & "\PAR", "HFTPAR52.dbf", "HR_FUTGrItems", "FoxPro 2.0", , False)
   End If
   
   Debug.Print "LinkMdbAdm: Tiempo: " & Format((CDbl(Now) - Tm) / TimeSerial(0, 0, 1), NUMFMT) & " [s]"
   
#End If
   
End Sub
Public Sub UnLinkMdbAdm_Old()

   Call ExecSQL(DbMain, "Drop Table " & "CodActiv")
   Call ExecSQL(DbMain, "Drop Table " & "Empresas")
   Call ExecSQL(DbMain, "Drop Table " & "EmpresasAno")
   Call ExecSQL(DbMain, "Drop Table " & "Equivalencia")
   Call ExecSQL(DbMain, "Drop Table " & "Impuestos")
   Call ExecSQL(DbMain, "Drop Table " & "Monedas")
   Call ExecSQL(DbMain, "Drop Table " & "Param")
   Call ExecSQL(DbMain, "Drop Table " & "PlanAvanzado")
   Call ExecSQL(DbMain, "Drop Table " & "PlanBasico")
   Call ExecSQL(DbMain, "Drop Table " & "PlanIntermedio")
   Call ExecSQL(DbMain, "Drop Table " & "Regiones")
   Call ExecSQL(DbMain, "Drop Table " & "Timbraje")
   Call ExecSQL(DbMain, "Drop Table " & "TipoValor")
   Call ExecSQL(DbMain, "Drop Table " & "Usuarios")
   Call ExecSQL(DbMain, "Drop Table " & "UsuarioEmpresa")
   Call ExecSQL(DbMain, "Drop Table " & "Perfiles")
   Call ExecSQL(DbMain, "Drop Table " & "IPC")
  
End Sub
Public Sub FillCbClasifEnt(Cb As ComboBox, Optional ByVal TipoEnt As Integer = ENT_CLIENTE)
   Dim i As Integer
   
   For i = ENT_CLIENTE To ENT_OTRO
      Cb.AddItem gClasifEnt(i)
      Cb.ItemData(Cb.NewIndex) = i
      
   Next i
   Cb.ListIndex = TipoEnt
   
End Sub
Public Sub ConfigNiveles()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Niv As String
   Dim TotNiv As Integer, IniNiv As Integer
   Dim i As Integer, CodNiv As Integer, LenNiv As Integer

   'NIVELES DE CUENTAS
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='NIVELES'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   TotNiv = vFld(Rs("Valor"))
   Call CloseRs(Rs)
   
   Niv = ""
   For i = 1 To TotNiv
      Niv = "," & "'DIGNIV" & i & "'" & Niv
   Next

   Niv = Mid(Niv, 2)
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo IN (" & Niv & ") "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   IniNiv = 1
   gFmtCodigoCta = ""
   Do While Rs.EOF = False
      CodNiv = vFld(Rs("Codigo"))
      LenNiv = vFld(Rs("Valor"))
   
      If LenNiv <> 0 Then
         gNiveles.nNiveles = CodNiv
         gNiveles.Inicio(CodNiv) = IniNiv
         gNiveles.Largo(CodNiv) = LenNiv
         IniNiv = IniNiv + LenNiv
         gLastNivel = CodNiv
         gFmtCodigoCta = gFmtCodigoCta & "-" & String(LenNiv, "0")
      Else
         Exit Sub
          
      End If
      
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   gFmtCodigoCta = Mid(gFmtCodigoCta, 2)
   
   'Para el plan IFRS
   gLastNivelIFRS = IFRS_MAXNIVEL
   gNivelesIFRS.nNiveles = 4
   gNivelesIFRS.Inicio(1) = 1
   gNivelesIFRS.Largo(1) = 1
   gNivelesIFRS.Inicio(2) = 2
   gNivelesIFRS.Largo(2) = 2
   gNivelesIFRS.Inicio(3) = 4
   gNivelesIFRS.Largo(3) = 2
   gNivelesIFRS.Inicio(4) = 6
   gNivelesIFRS.Largo(4) = 2
   
   
   gFmtCodigoIFRS = "0-00-00-00"

End Sub
Public Sub Foliacion()
   Dim Q1 As String
   Dim Rs As Recordset
   
   gFoliacion.Estado = EF_NOEXISTE
   
   gFoliacion.UltImpreso = 0
   gFoliacion.FUltImpreso = 0
   
   gFoliacion.UltTimbrado = 0
   gFoliacion.FUltTimbrado = 0
   
   gFoliacion.UltUsado = 0
   gFoliacion.FUltUsado = 0
   
   Q1 = "SELECT UltImpreso, FUltImpreso, UltTimbrado, FUltTimbrado, UltUsado, FUltUsado "
   Q1 = Q1 & " FROM Timbraje WHERE idEmpresa=" & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      gFoliacion.UltImpreso = vFld(Rs("UltImpreso"))
      gFoliacion.FUltImpreso = vFld(Rs("FUltImpreso"))
      
      gFoliacion.UltTimbrado = vFld(Rs("UltTimbrado"))
      gFoliacion.FUltTimbrado = vFld(Rs("FUltTimbrado"))
      
      gFoliacion.UltUsado = vFld(Rs("UltUsado"))
      gFoliacion.FUltUsado = vFld(Rs("FUltUsado"))
      
      gFoliacion.Estado = EF_EXISTE
      
   End If
   Call CloseRs(Rs)
End Sub
Public Sub FolioEncabEmpresa(EncabezadoEmp As Boolean, Orientacion As Integer)
   Dim Nombres(7) As String
   Dim i As Integer
   
   If EncabezadoEmp Then
   
      For i = 0 To UBound(Nombres)
         Nombres(i) = " "
      Next i
      
      gPrtLibros.TabNombres = FillMembreteEmp(Nombres)
            
      gPrtLibros.PrintNumPag = True
      gPrtLibros.PrintFecha = ChkNoPrtFecha
      
   Else
      'es foliado
      For i = 0 To UBound(Nombres)
         Nombres(i) = " "              'para que haga como que los imprime, por el espacio en el encabezado
      Next i
      
      If gEmpresa.RepConjunta = False Then
         Nombres(6) = ""
         Nombres(7) = ""
      End If
      
      gPrtLibros.PrintNumPag = False
      gPrtLibros.PrintFecha = False
   
   End If
   
   gPrtLibros.Nombres = Nombres
   'Printer.Orientation = Orientacion
   
End Sub
Public Sub FillCbCuentas(Cb As ComboBox, Optional ByVal LastNivel As Boolean = False)
   Dim Q1 As String
   Dim Rs As Recordset
   
    'LLenar combo
   Call AddItem(Cb, "(todas)", -1)
   Q1 = "SELECT Descripcion, Codigo, idCuenta FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   If LastNivel Then
      Q1 = Q1 & " AND Nivel = " & gLastNivel
   End If
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   Do While Rs.EOF = False
      Call CbAddItem(Cb, Format(vFld(Rs("Codigo")), gFmtCodigoCta) & " - " & FCase(vFld(Rs("Descripcion"), True)), vFld(Rs("idCuenta")))
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Cb.ListIndex = 0
         
End Sub

Public Function GetIdTipoLib(ByVal CodTipoLib As String) As Integer
   Dim i As Integer
   
   GetIdTipoLib = 0
   
   For i = 1 To UBound(gTipoLibCod)
      If gTipoLibCod(i) = CodTipoLib Then
         GetIdTipoLib = i
         Exit For
      End If
   Next i
   
End Function

Public Function GetDiminutivo(ByVal str As String) As String
   Dim i As Integer
   Dim Res As String
   Dim c As String
   
   For i = 1 To Len(str)
      c = Mid(str, i, 1)
      If UCase(c) = c And UCase(c) >= "A" And UCase(c) <= "Z" Then 'es mayúscula
         Res = Res & c
      End If
   Next i

   GetDiminutivo = Res
End Function
Public Function GetDiminutivoDoc(ByVal TipoLib As Integer, ByVal TipoDoc As Integer)
   Dim i As Integer
   
   GetDiminutivoDoc = ""
   
   For i = 0 To UBound(gTipoDoc)
      If gTipoDoc(i).Nombre = "" Then
         Exit Function
      End If
      
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).TipoDoc = TipoDoc Then
         
         If gTipoDoc(i).Diminutivo <> "" Then
            GetDiminutivoDoc = gTipoDoc(i).Diminutivo
         Else
            GetDiminutivoDoc = GetDiminutivo(gTipoDoc(i).Nombre)
         End If
         
         Exit Function
      End If
   Next i
   
End Function
'Public Function GenComprobante(ByVal StrIdDoc As String, ByVal TipoLib As Long, ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal idcomp As Long = 0) As Long
Public Function GenComprobante(ByVal StrIdDoc As String, ByVal TipoLib As Long, ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal idcomp As Long = 0, Optional ByVal CentrFull As Long = 0) As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tipo As Integer
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   Dim FirstIdDoc As String
   Dim Idx As Integer
   Dim IdCompNew As Long
   Dim Glosa As String
   Dim NomMes As String
   Dim Fecha As Long
   Dim MesActual As Integer
   Dim FirstDay As Long, LastDay As Long
   Dim FldArray(3) As AdvTbAddNew_t
   
   GenComprobante = 0
   
   MesActual = GetMesActual()   'debe haber un mes abierto
   
   If MesActual = 0 Then
      MsgBox "No hay mes abierto.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Q1 = "SELECT Count(*) "
   Q1 = Q1 & " FROM Documento LEFT JOIN MovDocumento "
   Q1 = Q1 & " ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento")
   Q1 = Q1 & " WHERE Documento.IdDoc IN (" & StrIdDoc & ")"
   Q1 = Q1 & " AND IdMovDoc IS NULL "
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      If vFld(Rs(0)) <> 0 Then    'las NCVs administrativas sólo se ingresan con los datos básicos, sin valores ni cuentas contables
         'MsgBox1 "No es posible realizar el proceso de Centralización debido a que hay documentos que no tienen movimientos contables asociados.", vbExclamation + vbOKOnly
         'MsgBox1 "ATENCIÓN: Hay documentos que no tienen movimientos contables asociados (NCV). Sólo se centralizarán aquellos documentos que cumplen con esta condición.", vbExclamation + vbOKOnly
         If MsgBox1("ATENCIÓN: Hay documentos de tipo NCV o NCC que tienen valor cero y que no se incluirán en la centralización." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Count(*) FROM MovDocumento "
   Q1 = Q1 & " WHERE IdDoc IN (" & StrIdDoc & ")"
   Q1 = Q1 & " AND IdCuenta = 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      If vFld(Rs(0)) <> 0 Then
         MsgBox1 "No es posible realizar el proceso de Centralización debido a que hay documentos que tienen cuentas en blanco.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   
   Call CloseRs(Rs)

   If idcomp = 0 Then    'viene en cero desde centralización en FrmCompraVenta
                         
'      Set Rs = DbMain.OpenRecordset("Comprobante")
'      Rs.AddNew
'
'      Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'      Rs.Fields("FechaCreacion") = CLng(Int(Now))
'
'      IdCompNew = Rs("IdComp")
'
'      Rs.Update
'      Rs.Close
'      Set Rs = Nothing
      
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
               
      IdCompNew = AdvTbAddNewMult(DbMain, "Comprobante", "IdComp", FldArray)
      
   End If
   
   If TipoLib = 0 Then
      
      Idx = InStr(StrIdDoc, ",")
      If Idx <= 0 Then
         FirstIdDoc = StrIdDoc
      Else
         FirstIdDoc = Left(StrIdDoc, Idx - 1)
      End If
   
      Q1 = "SELECT TipoLib FROM Documento WHERE IdDoc = " & FirstIdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         TipoLib = vFld(Rs("TipoLib"))
      End If
      
      Call CloseRs(Rs)
   End If
   
   If Mes >= 1 And Mes <= 12 Then
      NomMes = gNomMes(Mes)
   End If
   
   Tipo = TC_TRASPASO
   
   Select Case TipoLib
      Case LIB_VENTAS
         'Tipo = TC_INGRESO
         Glosa = "Centraliz. Ventas Mes de " & NomMes & " " & Ano
      Case LIB_COMPRAS
         'Tipo = TC_EGRESO
         Glosa = "Centraliz. Compras Mes de " & NomMes & " " & Ano
      Case LIB_REMU
         'Tipo = TC_EGRESO
         Glosa = "Centraliz. Remuneraciones Mes de " & NomMes & " " & Ano
      Case LIB_RETEN
         'Tipo = TC_EGRESO
         Glosa = "Centraliz. Retenciones Mes de " & NomMes & " " & Ano
      Case Else
         'Tipo = TC_INGRESO
   End Select
   
   Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovDocumento WHERE IdDoc IN (" & StrIdDoc & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
   End If
   
   Call CloseRs(Rs)
   
   
   If idcomp = 0 Then
      
      If MesActual = month(Now) Then
         Fecha = CLng(Int(DateSerial(gEmpresa.Ano, month(Now), Day(Now))))
      Else
         Fecha = DateSerial(gEmpresa.Ano, MesActual, 1)
      End If

      If gDtCompCent <> DTCOMPCENT_CURRDEF Then
         If gDtCompCent = DTCOMPCENT_DEFDAY And (month(Fecha) <> 2 And gDayDtCompCent > 0 And gDayDtCompCent <= 30) Or (month(Fecha) = 2 And gDayDtCompCent > 0 And gDayDtCompCent <= 28) Then
            Fecha = DateSerial(Year(Fecha), month(Fecha), gDayDtCompCent)
         Else
            Call FirstLastMonthDay(Fecha, FirstDay, LastDay)
            Fecha = DateSerial(Year(Fecha), month(Fecha), Day(LastDay))
         End If
      End If
      
      'actualizamos el encabezado
      Q1 = "UPDATE Comprobante SET "
      Q1 = Q1 & "  Fecha = " & Fecha
      Q1 = Q1 & ", Tipo = " & Tipo
      Q1 = Q1 & ", Estado = " & gEstadoNewComp
      Q1 = Q1 & ", Glosa = '" & Glosa & "'"
      Q1 = Q1 & ", TotalDebe = " & TotDebe
      Q1 = Q1 & ", TotalHaber = " & TotHaber
      Q1 = Q1 & ", ImpResumido = " & Abs(gImpResCent)
      Q1 = Q1 & ", TipoAjuste = " & TAJUSTE_AMBOS
      Q1 = Q1 & "  WHERE IdComp = " & IdCompNew
      Q1 = Q1 & "  AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
      Call ExecSQL(DbMain, Q1)
   
      idcomp = IdCompNew
   End If
   
   
   Q1 = "INSERT INTO MovComprobante (DeCentraliz, IdComp, IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT 1 as DeCentraliz, " & idcomp & " as IdComp, IIf(EsTotalDoc <> 0, MovDocumento.IdDoc ,0) As IdDoc, "
   Q1 = Q1 & " (MovDocumento.IdDoc * 100 +  MovDocumento.Orden) as Orden, IdCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg "
   Q1 = Q1 & "," & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano "
   Q1 = Q1 & " FROM MovDocumento "
   Q1 = Q1 & " WHERE MovDocumento.IdDoc IN(" & StrIdDoc & ") "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY MovDocumento.IdDoc, MovDocumento.Orden"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Documento SET IdCompCent = " & idcomp & ", Estado = " & ED_CENTRALIZADO & ", SaldoDoc = NULL WHERE IdDoc IN(" & StrIdDoc & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE MovDocumento SET IdCompCent = " & idcomp & " WHERE IdDoc IN(" & StrIdDoc & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   GenComprobante = idcomp
   
End Function

Public Function GenWhereCuentas(ByVal CodCuenta As String) As String
   Dim Idx As Integer
   Dim i As Integer
   
   Idx = 0
   For i = Len(CodCuenta) To 1 Step -1
      
      If Mid(CodCuenta, i, 1) = "-" Then
         Idx = i
      End If
      
      If Mid(CodCuenta, i, 1) <> "0" And Mid(CodCuenta, i, 1) <> "-" Then
         Exit For
      End If
   Next i
   
   If Idx > 0 Then
      CodCuenta = Left(CodCuenta, Idx - 1)
   End If
      
   CodCuenta = ReplaceStr(CodCuenta, "-", "")
   
   GenWhereCuentas = " Left(Cuentas.Codigo," & Len(CodCuenta) & ") = '" & CodCuenta & "'"

End Function

Public Function FillComboAno(Cb_Ano As ComboBox, Optional NAnosAtras As Integer = 5)
   Dim Ano As Long
   Dim i As Integer
   
   Ano = gEmpresa.Ano
   
   For i = NAnosAtras To 1 Step -1
      Cb_Ano.AddItem Ano
      Cb_Ano.ItemData(Cb_Ano.NewIndex) = Ano
      
      Ano = Ano - 1
      
   Next i
   
End Function

Public Function GenQueryPorNiveles(ByVal Nivel As Integer, ByVal Where As String, ByVal LibOficial As Boolean, Optional ByVal ClasCta As String = "", Optional ByVal Mensual As Boolean = False, Optional ByVal TipoDesglose As String = "", Optional ByVal WhDesglose As String = "", Optional ByVal Ano As Integer = 0, Optional ByVal ConOrderBy As Boolean = True) As String
   Dim Q1 As String
   Dim JoinComp As String
   Dim WhereEstado As String
   Dim N5 As Byte, N4 As Byte, N3 As Byte, N2 As Byte
   Dim SelMensual As String
   Dim SelMensual0 As String
   Dim SelDesglose As String
   Dim SelDesglose0 As String
   Dim GroupByMensual As String
   Dim GroupByDesglose As String
   Dim WhereEmpAnoComp As String, WhereEmpAnoCuentas As String
   
   N5 = gNiveles.nNiveles
   N4 = N5 - 1
   N3 = N4 - 1
   If N3 - 1 < 0 Then
      N2 = 0
   Else
      N2 = N3 - 1
   End If
   
   If Ano = 0 Then
      Ano = gEmpresa.Ano
   End If
   
   If LibOficial Then
      WhereEstado = " Comprobante.Estado=" & EC_APROBADO
      MsgBox1 "Dado que es Libro Oficial, sólo se seleccionarán los comprobantes APROBADOS.", vbInformation + vbOKOnly
   Else
      WhereEstado = " Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   End If
   
   'WhereFecha = " (Comprobante.Fecha BETWEEN " & GetTxDate(tx_Desde) & " AND " & GetTxDate(tx_Hasta) & ")"
'   JoinComp = " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp AND Comprobante.IdEmpresa = MovComprobante.IdEmpresa AND Comprobante.Ano = MovComprobante.Ano"
   JoinComp = " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp " & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   WhereEmpAnoComp = " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & Ano
   WhereEmpAnoCuentas = " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & Ano


   If Mensual Then
      SelMensual = ", " & SqlMonthLng("Comprobante.Fecha") & " As Mes"
      SelMensual0 = ", 0 As Mes"
      GroupByMensual = ", " & SqlMonthLng("Comprobante.Fecha")
   End If
   
   If TipoDesglose = "CCOSTO" Then
      SelDesglose = ", MovComprobante.IdCCosto as IdDesglose"
      SelDesglose0 = ", 0 as IdDesglose"
      GroupByDesglose = ", MovComprobante.IdCCosto "
      
   ElseIf TipoDesglose = "AREANEG" Then
      SelDesglose = ", MovComprobante.IdAreaNeg as IdDesglose"
      SelDesglose0 = ", 0 as IdDesglose"
      GroupByDesglose = ", MovComprobante.IdAreaNeg "
      
   Else
      WhDesglose = ""
            
   End If
   
   If WhDesglose <> "" Then
      WhDesglose = " AND " & WhDesglose
   End If

   'lista de cuentas de menor nivel
   Q1 = "SELECT 1 as IdQ, Cuentas.idCuenta, Codigo, Nivel, Descripcion, 0 as Debe, 0 As Haber, Cuentas.Clasificacion" & SelMensual0 & SelDesglose0
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE Nivel <= " & Nivel & WhereEmpAnoCuentas
   If ClasCta <> "" Then
      Q1 = Q1 & " AND Cuentas.Clasificacion IN(" & ClasCta & ")"
   End If
   
   Q1 = Q1 & " UNION"

   'lista de cuentas con nivel igual
   Q1 = Q1 & " SELECT 2 as IdQ, Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nivel, Cuentas.Descripcion, Sum(MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber, Cuentas.Clasificacion" & SelMensual & SelDesglose
   Q1 = Q1 & " FROM (Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
   Q1 = Q1 & "       AND MovComprobante.IdEmpresa = Cuentas.IdEmpresa )" & JoinComp
   'Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )" & JoinComp
   Q1 = Q1 & " WHERE Cuentas.Nivel <= " & Nivel & " AND " & Where & " AND " & WhereEstado & WhDesglose & WhereEmpAnoComp
   If ClasCta <> "" Then
      Q1 = Q1 & " AND Cuentas.Clasificacion IN(" & ClasCta & ")"
   End If
   Q1 = Q1 & " GROUP BY Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nivel, Cuentas.Descripcion, Cuentas.Clasificacion" & GroupByMensual & GroupByDesglose

   If Nivel < N5 Then

      Q1 = Q1 & " UNION"
      
      'suma de cuentas en que este nivel es el padre
      Q1 = Q1 & " SELECT 3 as IdQ, Cuentas_1.idCuenta, Cuentas_1.Codigo, Cuentas_1.Nivel, Cuentas_1.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber, Cuentas_1.Clasificacion" & SelMensual & SelDesglose
      Q1 = Q1 & " FROM ((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta"
'      Q1 = Q1 & "       AND MovComprobante.IdEmpresa = Cuentas.IdEmpresa AND MovComprobante.Ano = Cuentas.Ano )"
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta"
'      Q1 = Q1 & "       AND Cuentas.IdEmpresa = Cuentas_1.IdEmpresa AND Cuentas.Ano = Cuentas_1.Ano )" & JoinComp
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Cuentas_1") & " )" & JoinComp
      Q1 = Q1 & " WHERE Cuentas_1.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhDesglose & WhereEmpAnoComp
      If ClasCta <> "" Then
         Q1 = Q1 & " AND Cuentas_1.Clasificacion IN(" & ClasCta & ")"
      End If
      
      Q1 = Q1 & " GROUP BY Cuentas_1.idCuenta, Cuentas_1.Codigo, Cuentas_1.Nivel, Cuentas_1.Descripcion, Cuentas_1.Clasificacion" & GroupByMensual & GroupByDesglose
      
      If Nivel < N4 Then
         
         Q1 = Q1 & " UNION"
         
         'suma de cuentas en que este nivel es el abuelo
         Q1 = Q1 & " SELECT 4 as IdQ, Cuentas_2.idCuenta, Cuentas_2.Codigo, Cuentas_2.Nivel, Cuentas_2.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber, Cuentas_2.Clasificacion" & SelMensual & SelDesglose
         Q1 = Q1 & " FROM (((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
'         Q1 = Q1 & "       AND MovComprobante.IdEmpresa = Cuentas.IdEmpresa AND MovComprobante.Ano = Cuentas.Ano )"
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
         Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta"
'         Q1 = Q1 & "       AND Cuentas.IdEmpresa = Cuentas_1.IdEmpresa AND Cuentas.Ano = Cuentas_1.Ano )"
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Cuentas_1") & " )"
         Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta "
'         Q1 = Q1 & "       AND Cuentas_1.IdEmpresa = Cuentas_2.IdEmpresa AND Cuentas_1.Ano = Cuentas_2.Ano )" & JoinComp
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_2", "Cuentas_1") & " )" & JoinComp
         Q1 = Q1 & " WHERE Cuentas_2.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhDesglose & WhereEmpAnoComp
         If ClasCta <> "" Then
            Q1 = Q1 & " AND Cuentas_2.Clasificacion IN(" & ClasCta & ")"
         End If
         
         Q1 = Q1 & " GROUP BY Cuentas_2.idCuenta, Cuentas_2.Codigo, Cuentas_2.Nivel, Cuentas_2.Descripcion, Cuentas_2.Clasificacion" & GroupByMensual & GroupByDesglose
         
         If Nivel < N3 Then
            
            Q1 = Q1 & " UNION"
            
            'suma de cuentas en que este nivel es el bis-abuelo
            Q1 = Q1 & " SELECT 5 as IdQ, Cuentas_3.idCuenta, Cuentas_3.Codigo, Cuentas_3.Nivel, Cuentas_3.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber, Cuentas_3.Clasificacion" & SelMensual & SelDesglose
            Q1 = Q1 & " FROM ((((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta"
'            Q1 = Q1 & "       AND MovComprobante.IdEmpresa = Cuentas.IdEmpresa AND MovComprobante.Ano = Cuentas.Ano )"
            Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
            Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta "
'            Q1 = Q1 & "       AND Cuentas.IdEmpresa = Cuentas_1.IdEmpresa AND Cuentas.Ano = Cuentas_1.Ano) "
            Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Cuentas_1") & " )"
            Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta "
'            Q1 = Q1 & "       AND Cuentas_1.IdEmpresa = Cuentas_2.IdEmpresa AND Cuentas_1.Ano = Cuentas_2.Ano )"
            Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_2", "Cuentas_1") & " )"
            Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_3 ON Cuentas_2.idPadre = Cuentas_3.idCuenta "
'            Q1 = Q1 & "       AND Cuentas_2.IdEmpresa = Cuentas_3.IdEmpresa AND Cuentas_2.Ano = Cuentas_3.Ano )" & JoinComp
            Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_3", "Cuentas_2") & " )" & JoinComp
            
            Q1 = Q1 & " WHERE Cuentas_3.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhDesglose & WhereEmpAnoComp
            If ClasCta <> "" Then
               Q1 = Q1 & " AND Cuentas_3.Clasificacion IN(" & ClasCta & ")"
            End If
                        
            Q1 = Q1 & " GROUP BY Cuentas_3.idCuenta, Cuentas_3.Codigo, Cuentas_3.Nivel, Cuentas_3.Descripcion, Cuentas_3.Clasificacion" & GroupByMensual & GroupByDesglose
            
            If Nivel < N2 Then
            
               Q1 = Q1 & " UNION"
               
               'suma de cuentas en que este nivel es el tatara-abuelo
               Q1 = Q1 & " SELECT 6 as IdQ, Cuentas_4.idCuenta, Cuentas_4.Codigo, Cuentas_4.Nivel, Cuentas_4.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber, Cuentas_4.Clasificacion" & SelMensual & SelDesglose
               Q1 = Q1 & " FROM (((((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta) "
'               Q1 = Q1 & "       AND MovComprobante.IdEmpresa = Cuentas.IdEmpresa AND MovComprobante.Ano = Cuentas.Ano )"
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
               Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta "
'               Q1 = Q1 & "       AND Cuentas.IdEmpresa = Cuentas_1.IdEmpresa AND Cuentas.Ano = Cuentas_1.Ano )"
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Cuentas_1") & " )"
               Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta "
'               Q1 = Q1 & "       AND Cuentas_1.IdEmpresa = Cuentas_2.IdEmpresa AND Cuentas_1.Ano = Cuentas_2.Ano )"
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_2", "Cuentas_1") & " )"
               Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_3 ON Cuentas_2.idPadre = Cuentas_3.idCuenta "
'               Q1 = Q1 & "       AND Cuentas_2.IdEmpresa = Cuentas_3.IdEmpresa AND Cuentas_2.Ano = Cuentas_3.Ano )" & JoinComp
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_3", "Cuentas_2") & " )"
               Q1 = Q1 & " INNER JOIN Cuentas AS Cuentas_4 ON Cuentas_3.idPadre = Cuentas_4.idCuenta "
'               Q1 = Q1 & "       AND Cuentas_3.IdEmpresa = Cuentas_4.IdEmpresa AND Cuentas_3.Ano = Cuentas_4.Ano )" & JoinComp
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas_3", "Cuentas_4") & " )" & JoinComp
               Q1 = Q1 & " WHERE Cuentas_4.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhDesglose & WhereEmpAnoComp
               If ClasCta <> "" Then
                  Q1 = Q1 & " AND Cuentas_4.Clasificacion IN(" & ClasCta & ")"
               End If
               
               Q1 = Q1 & " GROUP BY Cuentas_4.idCuenta, Cuentas_4.Codigo, Cuentas_4.Nivel, Cuentas_4.Descripcion,Cuentas_4.Clasificacion" & GroupByMensual & GroupByDesglose
            End If
         
         End If
         
      End If
   
   End If
         
   If ConOrderBy Then
      Q1 = Q1 & " ORDER BY Codigo"
   End If

   GenQueryPorNiveles = Q1
   
End Function

Public Function GenQueryIFRSporNiveles(ByVal Nivel As Integer, ByVal Where As String, ByVal LibOficial As Boolean, Optional ByVal TipoInfoIFRS As Integer = 0, Optional ByVal Mensual As Boolean = False) As String
   Dim Q1 As String
   Dim JoinComp As String
   Dim WhereEstado As String
   Dim N5 As Byte, N4 As Byte, N3 As Byte, N2 As Byte
   Dim SelMensual As String
   Dim SelMensual0 As String
   Dim GroupByMensual As String
   Dim TblIFRS As String
   Dim TmpTbl As String
   Dim Fld As String, FldR As String
   Dim WhereEmpAnoComp As String, WhereEmpAnoCuentas As String
   
   N5 = gNiveles.nNiveles
   N4 = N5 - 1
   N3 = N4 - 1
   If N3 - 1 < 0 Then
      N2 = 0
   Else
      N2 = N3 - 1
   End If
         
   If LibOficial Then
      WhereEstado = " Comprobante.Estado=" & EC_APROBADO
      'MsgBox1 "Dado que es Libro Oficial, sólo se seleccionarán los comprobantes APROBADOS.", vbInformation + vbOKOnly
   Else
      WhereEstado = " Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   End If
      
   JoinComp = " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'   JoinComp = JoinComp & " AND Comprobante.IdEmpresa = MovComprobante.IdEmpresa AND Comprobante.Ano = MovComprobante.Ano "
   JoinComp = JoinComp & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   
   WhereEmpAnoComp = "AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   WhereEmpAnoCuentas = "AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano

   If Mensual Then
      SelMensual = ", " & SqlMonthLng("Comprobante.Fecha") & " As Mes"
      SelMensual0 = ", 0 As Mes"
      GroupByMensual = ", " & SqlMonthLng("Comprobante.Fecha")
   End If

   'lista de cuentas de menor nivel
   Q1 = "SELECT 1 as IdQ, IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion, 0 as Debe, 0 As Haber " & SelMensual0
   Q1 = Q1 & " FROM IFRS_PlanIFRS "
   Q1 = Q1 & " LEFT JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS "
   Q1 = Q1 & " WHERE IFRS_PlanIFRS.Nivel <= " & Nivel & WhereEmpAnoCuentas
   
   Q1 = Q1 & " UNION"

   'lista de cuentas con nivel igual
   Q1 = Q1 & " SELECT 2 as IdQ, IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion, Sum(MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber " & SelMensual
   Q1 = Q1 & " FROM ((IFRS_PlanIFRS "
   Q1 = Q1 & " INNER JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
   Q1 = Q1 & " INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
'   Q1 = Q1 & "       AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa AND Cuentas.Ano = MovComprobante.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & JoinComp
   Q1 = Q1 & " WHERE IFRS_PlanIFRS.Nivel <= " & Nivel & " AND " & Where & " AND " & WhereEstado & WhereEmpAnoCuentas
   Q1 = Q1 & " GROUP BY IFRS_PlanIFRS.idCuenta, IFRS_PlanIFRS.Codigo, IFRS_PlanIFRS.Nivel, IFRS_PlanIFRS.Descripcion " & GroupByMensual

   If Nivel < N5 Then

      Q1 = Q1 & " UNION"
      
      'suma de cuentas en que este nivel es el padre
      Q1 = Q1 & " SELECT 3 as IdQ, IFRS_PlanIFRS_1.idCuenta, IFRS_PlanIFRS_1.Codigo, IFRS_PlanIFRS_1.Nivel, IFRS_PlanIFRS_1.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber " & SelMensual
      Q1 = Q1 & " FROM (((IFRS_PlanIFRS "
      Q1 = Q1 & " INNER JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
      Q1 = Q1 & " INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
'      Q1 = Q1 & "       AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa AND Cuentas.Ano = MovComprobante.Ano )"
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_1 ON IFRS_PlanIFRS.idPadre = IFRS_PlanIFRS_1.idCuenta)" & JoinComp
      Q1 = Q1 & " WHERE IFRS_PlanIFRS_1.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhereEmpAnoCuentas
      
      Q1 = Q1 & " GROUP BY IFRS_PlanIFRS_1.idCuenta, IFRS_PlanIFRS_1.Codigo, IFRS_PlanIFRS_1.Nivel, IFRS_PlanIFRS_1.Descripcion " & GroupByMensual
      
      If Nivel < N4 Then
         
         Q1 = Q1 & " UNION"
         
         'suma de cuentas en que este nivel es el abuelo
         Q1 = Q1 & " SELECT 4 as IdQ, IFRS_PlanIFRS_2.idCuenta, IFRS_PlanIFRS_2.Codigo, IFRS_PlanIFRS_2.Nivel, IFRS_PlanIFRS_2.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber " & SelMensual
         Q1 = Q1 & " FROM ((((IFRS_PlanIFRS "
         Q1 = Q1 & " INNER JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
         Q1 = Q1 & " INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta"
'         Q1 = Q1 & "       AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa AND Cuentas.Ano = MovComprobante.Ano) "
         Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
         Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_1 ON IFRS_PlanIFRS.idPadre = IFRS_PlanIFRS_1.idCuenta ) "
         Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_2 ON IFRS_PlanIFRS_1.idPadre = IFRS_PlanIFRS_2.idCuenta)" & JoinComp
         Q1 = Q1 & " WHERE IFRS_PlanIFRS_2.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhereEmpAnoCuentas
         
         Q1 = Q1 & " GROUP BY IFRS_PlanIFRS_2.idCuenta, IFRS_PlanIFRS_2.Codigo, IFRS_PlanIFRS_2.Nivel, IFRS_PlanIFRS_2.Descripcion " & GroupByMensual
         
         If Nivel < N3 Then
            
            Q1 = Q1 & " UNION"
            
            'suma de cuentas en que este nivel es el bis-abuelo
            Q1 = Q1 & " SELECT 5 as IdQ, IFRS_PlanIFRS_3.idCuenta, IFRS_PlanIFRS_3.Codigo, IFRS_PlanIFRS_3.Nivel, IFRS_PlanIFRS_3.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber " & SelMensual
            Q1 = Q1 & " FROM (((((IFRS_PlanIFRS "
            Q1 = Q1 & " INNER JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
            Q1 = Q1 & " INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
'            Q1 = Q1 & "       AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa AND Cuentas.Ano = MovComprobante.Ano )"
            Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
            Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_1 ON IFRS_PlanIFRS.idPadre = IFRS_PlanIFRS_1.idCuenta) "
            Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_2 ON IFRS_PlanIFRS_1.idPadre = IFRS_PlanIFRS_2.idCuenta)"
            Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_3 ON IFRS_PlanIFRS_2.idPadre = IFRS_PlanIFRS_3.idCuenta)" & JoinComp
            Q1 = Q1 & " WHERE IFRS_PlanIFRS_3.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhereEmpAnoCuentas
                         
            Q1 = Q1 & " GROUP BY IFRS_PlanIFRS_3.idCuenta, IFRS_PlanIFRS_3.Codigo, IFRS_PlanIFRS_3.Nivel, IFRS_PlanIFRS_3.Descripcion " & GroupByMensual
            
            If Nivel < N2 Then
            
               Q1 = Q1 & " UNION"
               
               'suma de cuentas en que este nivel es el tatara-abuelo
               Q1 = Q1 & " SELECT 6 as IdQ, IFRS_PlanIFRS_4.idCuenta, IFRS_PlanIFRS_4.Codigo, IFRS_PlanIFRS_4.Nivel, IFRS_PlanIFRS_4.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber " & SelMensual
               Q1 = Q1 & " FROM ((((((IFRS_PlanIFRS "
               Q1 = Q1 & " INNER JOIN Cuentas ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS) "
               Q1 = Q1 & " INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
'               Q1 = Q1 & "       AND Cuentas.IdEmpresa = MovComprobante.IdEmpresa AND Cuentas.Ano = MovComprobante.Ano )"
               Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
               Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_1 ON IFRS_PlanIFRS.idPadre = IFRS_PlanIFRS_1.idCuenta) "
               Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_2 ON IFRS_PlanIFRS_1.idPadre = IFRS_PlanIFRS_2.idCuenta)"
               Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_3 ON IFRS_PlanIFRS_2.idPadre = IFRS_PlanIFRS_3.idCuenta)"
               Q1 = Q1 & " INNER JOIN IFRS_PlanIFRS AS IFRS_PlanIFRS_4 ON IFRS_PlanIFRS_3.idPadre = IFRS_PlanIFRS_4.idCuenta)" & JoinComp
               Q1 = Q1 & " WHERE IFRS_PlanIFRS_4.Nivel = " & Nivel & " AND " & Where & " AND " & WhereEstado & WhereEmpAnoCuentas
               
               Q1 = Q1 & " GROUP BY IFRS_PlanIFRS_4.idCuenta, IFRS_PlanIFRS_4.Codigo, IFRS_PlanIFRS_4.Nivel, IFRS_PlanIFRS_4.Descripcion " & GroupByMensual
            End If
         
         End If
         
      End If
   
   End If
         
   Q1 = Q1 & " ORDER BY Codigo"

   GenQueryIFRSporNiveles = Q1
   
End Function

Public Function GetIdEntidad(ByVal Rut As String, Nombre As String, NotValidRut As Boolean) As Long
   Dim Q1 As String
   Dim Rs As Recordset
   
   Rut = Trim(Rut)
   
   Q1 = "SELECT IdEntidad, Nombre, Rut, NotValidRut FROM Entidades "
   Q1 = Q1 & " WHERE (Rut = '" & vFmtCID(Rut) & "' OR Rut = '" & Rut & "')"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY Rut Desc"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      If vFld(Rs("NotValidRut")) <> 0 Then   'es RUT extranjero
         
         If vFld(Rs("Rut")) = Rut Then
            GetIdEntidad = vFld(Rs("IdEntidad"))
            Nombre = vFld(Rs("Nombre"))
            NotValidRut = vFld(Rs("NotValidRut"))
         ElseIf vFld(Rs("Rut")) = vFmtCID(Rut) Then
            GetIdEntidad = vFld(Rs("IdEntidad"))
            Nombre = vFld(Rs("Nombre"))
            NotValidRut = vFld(Rs("NotValidRut"))
         Else
            GetIdEntidad = 0
            Nombre = ""
         End If
      
      Else
         
         If vFld(Rs("Rut")) = vFmtCID(Rut) Then
            GetIdEntidad = vFld(Rs("IdEntidad"))
            Nombre = vFld(Rs("Nombre"))
            NotValidRut = vFld(Rs("NotValidRut"))
         Else
            GetIdEntidad = 0
            Nombre = ""
         End If
      End If
      
   Else
      GetIdEntidad = 0
      Nombre = ""
   End If
   
   Call CloseRs(Rs)
   
End Function


Public Function AddEntidad(ByVal Rut As String, ByVal RazonSocial As String, IdEntidad As Long, Optional ByVal Clasif As Integer = 0, Optional ByVal NotValidRut As Boolean = False) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Codigo As String
   Dim FldArray(3) As AdvTbAddNew_t

'   Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
'   Rs.AddNew
'
'   IdEntidad = Rs("idEntidad")
'
'   Rs("NotValidRut") = 0
'   Rs("RUT") = vFmtCID(Rut)
'
'   Rs.Update
'   Rs.Close
   
   If NotValidRut Then
      Codigo = Left(Rut, 15)
   Else
      Codigo = vFmtRut(Rut)
   End If
   
   FldArray(0).FldName = "NotValidRut"
   FldArray(0).FldValue = IIf(NotValidRut, 1, 0)
   FldArray(0).FldIsNum = True
   
   FldArray(1).FldName = "RUT"
   FldArray(1).FldValue = Left(vFmtCID(Rut, Not NotValidRut), 12)
   FldArray(1).FldIsNum = False
         
   FldArray(2).FldName = "IdEmpresa"
   FldArray(2).FldValue = gEmpresa.id
   FldArray(2).FldIsNum = True
                     
   FldArray(3).FldName = "Codigo"
   FldArray(3).FldValue = Codigo
   FldArray(3).FldIsNum = False
                     
   IdEntidad = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)

   
   If ERR Then
      IdEntidad = 0
      AddEntidad = False
      Exit Function
   End If
   
      
   Q1 = "UPDATE Entidades SET "
   Q1 = Q1 & "  Nombre='" & ParaSQL(Left(RazonSocial, 100)) & "'"
   Q1 = Q1 & ", Codigo='" & Codigo & "'"
   If Clasif >= 0 Then
      Q1 = Q1 & ", Clasif" & Clasif & " = 1"
   End If
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Call ExecSQL(DbMain, Q1)
   
   AddEntidad = True
   
End Function


Public Function GetCorrelativoComp(ByVal idcomp As Long) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Correlativo FROM Comprobante WHERE IdComp = " & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      GetCorrelativoComp = vFld(Rs("Correlativo"))
   Else
      GetCorrelativoComp = 0
   End If
   
   Call CloseRs(Rs)
      
End Function
Public Function AbrirMes(ByVal Mes As Integer) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim EstadoMes As Integer
   Dim Impreso As Boolean
   
   AbrirMes = False
   
    'SF  14663216
   Q1 = ""
   Q1 = "SELECT count(Mes) as numMes FROM EstadoMes "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " And Mes = " & Mes
   
   Set Rs = OpenRs(DbMain, Q1)
     
   If Rs.EOF = False Then
      
      If vFld(Rs("numMes")) = 0 Then
        Q1 = ""
        Q1 = "INSERT INTO EstadoMes "
        Q1 = Q1 & " (Mes, IdEmpresa, Ano, Estado, FechaApertura, FechaCierre) "
        Q1 = Q1 & " VALUES(" & Mes & "," & gEmpresa.id & ", " & gEmpresa.Ano & ", " & EM_CERRADO & ", 0, 0 )"
        Call ExecSQL(DbMain, Q1)
               
        AddLog ("Mes insertado = " & Mes)
       End If
    End If
   
   Call CloseRs(Rs)
   AddLog ("Cerramos Rs")
   'SF  14663216
   
   EstadoMes = GetEstadoMes(Mes)
   Impreso = LibMensualesImpresos(Mes, False)
   
   'If Mes > GetUltimoMesConMovs() + 1 Then
   '   MsgBox1 "No es posible abrir este mes.", vbExclamation + vbOKOnly
   '   Exit Function
   'End If
   
   Select Case EstadoMes
      Case EM_ABIERTO
         MsgBox1 "Este mes ya está abierto.", vbExclamation + vbOKOnly
         Exit Function
      Case EM_CERRADO
         If Impreso = True Then
            If MsgBox1("Este mes ya fue impreso en forma oficial." & vbNewLine & vbNewLine & "¿Está seguro que desea abrirlo?", vbQuestion + vbYesNo) = vbNo Then
               Exit Function
            End If
         End If
      Case EM_NOEXISTE
         MsgBox1 "Este mes no existe en la base de datos.", vbExclamation + vbOKOnly
         Exit Function
      Case EM_ERRONEO
         MsgBox1 "Este mes no está cuadrado. Se abrirá para que lo cuadre.", vbExclamation + vbOKOnly
   End Select
      
   Q1 = "UPDATE EstadoMes SET Estado = " & EM_ABIERTO & ", FechaApertura = " & CLng(Int(Now)) & " WHERE Mes = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
      
   AbrirMes = True
   
End Function
Public Function CerrarMes(ByVal Mes As Integer) As Boolean
   Dim Q1 As String
   
   CerrarMes = False
   
   If Not ValidaCierreMes(Mes) Then
      Exit Function
   End If
   
   Q1 = "UPDATE EstadoMes SET Estado = " & EM_CERRADO & ", FechaCierre = " & CLng(Int(Now)) & " WHERE Mes = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
      
   CerrarMes = True

End Function

Public Function ValidaCierreMes(ByVal Mes As Integer) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim EstadoMes As Integer
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim TmpTbl As String
   
   ValidaCierreMes = True
     
   EstadoMes = GetEstadoMes(Mes)
   If EstadoMes <> EM_ABIERTO Then
      MsgBox1 "Este mes no está abierto.", vbExclamation + vbOKOnly
      ValidaCierreMes = False
      Exit Function
   End If
   
   If LibMensualesImpresos(Mes, True) = False Then
      ValidaCierreMes = False
      Exit Function
   End If
   
   'vemos si no quedan comprobantes pendientes
   Q1 = "SELECT IdComp FROM Comprobante WHERE Estado = " & EC_PENDIENTE & " AND " & SqlMonthLng("Fecha") & " = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      If MsgBox1("El mes de " & gNomMes(Mes) & " tiene comprobantes en estado 'Pendiente'." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         ValidaCierreMes = False
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   
   Call CloseRs(Rs)
   
   'vemos si quedan documentos sin centralizar
   Q1 = "SELECT IdDoc FROM Documento WHERE Estado IN (" & ED_PENDIENTE & "," & ED_APROBADO & ") AND " & SqlMonthLng("FEmision") & " = " & Mes & " AND " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano & " AND TipoLib IN (" & LIB_VENTAS & "," & LIB_COMPRAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      If MsgBox1("El mes de " & gNomMes(Mes) & " tiene documentos que aún no han sido centralizados." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         ValidaCierreMes = False
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   
   Call CloseRs(Rs)
   
   'vemos si el mes está cuadrado
   Q1 = "SELECT Sum(TotalDebe) as TotDebe, Sum(TotalHaber) as TotHaber FROM Comprobante WHERE " & SqlMonthLng("Fecha") & " = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      If vFld(Rs("TotDebe")) <> vFld(Rs("TotHaber")) Then
         MsgBox1 "El mes de " & gNomMes(Mes) & " no está cuadrado. No es posible cerrar el mes." & vbNewLine & vbNewLine & "Revise la lista de comprobantes para ver cuáles tienen estado ERRONEO.", vbExclamation + vbOKOnly
         ValidaCierreMes = False
      End If
   End If
   
   Call CloseRs(Rs)
   
   'marcamos como erróneos los comprobantes no cuadrados
   If ValidaCierreMes = False Then
   
      TmpTbl = DbGenTmpName2(gDbType, "TmpCompErr")
      
      Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
      
      Q1 = " SELECT MovComprobante.IdComp, Sum(MovComprobante.Debe) as TDebe, Sum(MovComprobante.Haber) As THaber, " & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano"
      Q1 = Q1 & " INTO  " & TmpTbl
      Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE " & SqlMonthLng("Comprobante.Fecha") & " = " & Mes
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " GROUP BY MovComprobante.IdComp "
      Q1 = Q1 & " HAVING Sum(MovComprobante.Debe) <> Sum(MovComprobante.Haber)"
      Call ExecSQL(DbMain, Q1)
   
'      Q1 = "UPDATE Comprobante INNER JOIN TmpCompErroneos ON Comprobante.IdComp = TmpCompErroneos.IdComp "
'      Q1 = Q1 & " AND Comprobante.IdEmpresa = TmpCompErroneos.IdEmpresa AND Comprobante.Ano = TmpCompErroneos.Ano"
'      Q1 = Q1 & " SET Comprobante.Estado = " & EC_ERRONEO
'      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'      Call ExecSQL(DbMain, Q1)
      
      Tbl = " Comprobante "
      sFrom = " Comprobante INNER JOIN " & TmpTbl & " ON Comprobante.IdComp =  " & TmpTbl & ".IdComp "
      sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", TmpTbl)
      sSet = " Comprobante.Estado = " & EC_ERRONEO
      sWhere = " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano

      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
      
      Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
      
   End If
   
End Function
Public Function GetEstadoMes(ByVal Mes As Integer) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Estado FROM EstadoMes WHERE Mes = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      GetEstadoMes = vFld(Rs("Estado"))
   Else
      GetEstadoMes = EM_NOEXISTE
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Sub LlenarTablaMeses(ByVal Tbl As String)
   Dim Q1 As String
   Dim i As Integer
   
   If Tbl = "" Then
      Tbl = "EstadoMes"
   End If
   
   For i = 1 To 12
      Q1 = "INSERT INTO " & Tbl
      Q1 = Q1 & " (Mes, IdEmpresa, Ano, Estado, FechaApertura, FechaCierre) "
      Q1 = Q1 & " VALUES(" & i & "," & gEmpresa.id & ", " & gEmpresa.Ano & ", " & EM_CERRADO & ", 0, 0 )"
      Call ExecSQL(DbMain, Q1)
   Next i

End Sub
Public Function GetImpresoMes_Old(ByVal Mes As Integer) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Impreso FROM EstadoMes WHERE Mes = " & Mes
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      GetImpresoMes_Old = (vFld(Rs("Impreso")) <> 0)
   Else
      GetImpresoMes_Old = 0
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function GetAtribCuenta(ByVal IdCuenta As Long, ByVal Atrib As Integer) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetAtribCuenta = False
   
   If Atrib < 1 Or Atrib > MAX_ATRIB Then
      Exit Function
   End If
   
   Q1 = "SELECT Atrib" & Atrib & " FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      GetAtribCuenta = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function GetClasCuenta(ByVal IdCuenta As Long) As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetClasCuenta = False
   
   Q1 = "SELECT Clasificacion FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      GetClasCuenta = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function GetPathCuenta(ByVal IdCuenta As Long) As String
   Dim Path As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdPadre As Long
   
   IdPadre = IdCuenta
   
   Do
      Q1 = "SELECT Descripcion, IdPadre FROM Cuentas WHERE IdCuenta = " & IdPadre
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         Path = FCase(vFld(Rs("Descripcion"), True)) & " / " & Path
         IdPadre = vFld(Rs("IdPadre"))
      Else
         Exit Do
      End If
      
      Call CloseRs(Rs)
      
   Loop
   
   Call CloseRs(Rs)
   
   If Path <> "" Then
      Path = Left(Path, Len(Path) - 3)
   End If
   
   GetPathCuenta = Path
   
End Function

Public Function ValidaIngresoComp(Optional ByVal ModificarComp As Boolean = False, Optional ByVal EliminarComp As Boolean = False) As Boolean
   Dim MesActual As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   ValidaIngresoComp = False
   
   If gAppCode.Demo Then
      If Not ModificarComp And Not EliminarComp Then   'sólo new
         Q1 = "SELECT Count(*) FROM Comprobante "
         Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
      
         If Not Rs.EOF Then
            If vFld(Rs(0)) >= MAX_COMPDEMO Then
               MsgBox1 "Ha superado la cantidad de comprobantes de la versión DEMO.", vbExclamation
               Call CloseRs(Rs)
               Exit Function
            End If
         End If
         Call CloseRs(Rs)
      End If
   End If

      
   If gTipoCorrComp <= 0 Then
      MsgBox1 "No es posible crear comprobantes sin antes definir la configuración del correlativo.", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   MesActual = GetMesActual()
   
   If MesActual = 0 Then
      MsgBox1 "No es posible ingresar/modificar/eliminar comprobantes. No hay mes abierto.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If Not ModificarComp Then   'nuevo comp o eliminar comp
      If MesActual < GetUltimoMesConComps() And (gTipoCorrComp = TCC_UNICO And (gPerCorrComp = TCC_ANUAL Or gPerCorrComp = TCC_CONTINUO)) Then
         MsgBox1 "No es posible ingresar/eliminar comprobantes en el mes actual (" & gNomMes(MesActual) & ") porque la configuración de correlativo seleccionada para la empresa no lo permite.", vbExclamation + vbOKOnly
         Exit Function
      End If
   End If
   
   ValidaIngresoComp = True

End Function

Public Function ValidaIngresoDoc() As Boolean
   Dim MesActual As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   ValidaIngresoDoc = False
   
   If gAppCode.Demo Then
      Q1 = "SELECT Count(*) FROM Documento "
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
   
      If Not Rs.EOF Then
         If vFld(Rs(0)) >= MAX_DOCDEMO + 1 Then
            MsgBox1 "Ha superado la cantidad de documentos de la versión DEMO.", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
      Call CloseRs(Rs)
   End If
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   MesActual = GetMesActual()
   
   If MesActual = 0 Then
      MsgBox1 "No es posible ingresar documentos. No hay mes abierto.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   ValidaIngresoDoc = True

End Function

Public Function GetMesActual() As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Max(Mes) As MaxMes FROM EstadoMes WHERE Estado = " & EM_ABIERTO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      GetMesActual = vFld(Rs("MaxMes"))
   Else
      GetMesActual = 0
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function GetUltimoMesConComps() As Integer
   Dim Rs As Recordset
   Dim MaxFecha As Long
   Dim Q1 As String
   
   MaxFecha = 0
   
   Q1 = "SELECT Max(Fecha) FROM Comprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MaxFecha = vFld(Rs(0))
   End If
      
   Call CloseRs(Rs)
   
   If MaxFecha > 0 Then
      GetUltimoMesConComps = month(MaxFecha)
   Else
      GetUltimoMesConComps = 1  'partimos con enero
   End If
      
End Function


Public Function GetUltimoMesConMovs(Optional ByVal Msg As Boolean = False) As Integer
   Dim Rs As Recordset
   Dim MaxFechaComp As Long
   Dim MaxFechaDoc As Long
   Dim MaxFecha As Long
   Dim Q1 As String
   
   MaxFecha = 0
   MaxFechaComp = 0
   MaxFechaDoc = 0
   
   Q1 = "SELECT Max(Fecha) FROM Comprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MaxFechaComp = vFld(Rs(0))
      
      If Year(MaxFechaComp) > gEmpresa.Ano Then
         
         If Msg Then
            MsgBox1 "ATENCIÓN:" & vbNewLine & vbNewLine & "Hay comprobantes cuya fecha de emisión no corresponde al año actual. Verifíquelo en la lista de comprobantes y realice las modificaciones correspondientes.", vbExclamation
         End If
         
         Call CloseRs(Rs)
         
         Set Rs = OpenRs(DbMain, "SELECT Max(Fecha) FROM Comprobante WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND " & SqlYearLng("Fecha") & " = " & gEmpresa.Ano)
         If Rs.EOF = False Then
            MaxFechaComp = vFld(Rs(0))
         End If
      End If
      
   End If
      
   Call CloseRs(Rs)
   
   Set Rs = OpenRs(DbMain, "SELECT Max(FEmision) FROM Documento WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      If vFld(Rs(0)) > MaxFecha Then
         MaxFechaDoc = vFld(Rs(0))
         
         If Year(MaxFechaDoc) > gEmpresa.Ano Then
            
            If Msg Then
               MsgBox1 "ATENCIÓN:" & vbNewLine & vbNewLine & "Hay documentos cuya fecha de emisión es posterior al 31 diciembre " & gEmpresa.Ano & ". Verifíquelo en la lista de documentos y realice las modificaciones correspondientes.", vbExclamation
            End If
            
            Call CloseRs(Rs)
            
            Set Rs = OpenRs(DbMain, "SELECT Max(FEmision) FROM Documento WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND " & SqlYearLng("FEmision") & " <= " & gEmpresa.Ano)
            If Rs.EOF = False Then
               MaxFechaDoc = vFld(Rs(0))
            End If
            
         End If
         
      End If
   End If
      
   Call CloseRs(Rs)
   
   If MaxFechaDoc > MaxFechaComp Then
      MaxFecha = MaxFechaDoc
   Else
      MaxFecha = MaxFechaComp
   End If
   
   If MaxFecha > 0 Then
      GetUltimoMesConMovs = month(MaxFecha)
   Else
      GetUltimoMesConMovs = 1  'partimos con enero
   End If
      
End Function

Public Function GetNombreUsuario(ByVal IdUsuario As Long) As String
   Dim Rs As Recordset
   
   Set Rs = OpenRs(DbMain, "SELECT Usuario FROM Usuarios WHERE IdUsuario=" & IdUsuario)
   
   If Rs.EOF = False Then
      GetNombreUsuario = vFld(Rs("Usuario"), True)
   Else
      GetNombreUsuario = ""
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function CompAperturaTribTieneMovs() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim idcomp As Long, NMov As Long
   

   CompAperturaTribTieneMovs = False
   
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo =" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   idcomp = 0
   If Rs.EOF = False Then
      idcomp = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   If idcomp = 0 Then
      Exit Function
   End If
   
   Q1 = "SELECT Count(*) FROM MovComprobante WHERE IdComp = " & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   NMov = 0
   If Rs.EOF = False Then
      NMov = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   If NMov > 0 Then
      CompAperturaTribTieneMovs = True
   End If
   
End Function

Public Function EsCuentaActFijo(ByVal IdCuenta As Long) As Boolean

   EsCuentaActFijo = False
      
   If IdCuenta <> 0 Then
      EsCuentaActFijo = (GetAtribCuenta(IdCuenta, ATRIB_ACTIVOFIJO) <> 0)
   End If
   
End Function

Public Sub FillNivel(Cb As ComboBox, Optional Sel As Integer = -1)
   Dim i As Integer
   
   For i = 1 To gNiveles.nNiveles
      Cb.AddItem i
      Cb.ItemData(Cb.NewIndex) = i
      
      If i = Sel Then
         Cb.ListIndex = Cb.NewIndex
      End If
      
   Next i
   
   If Sel = -1 Then
      Cb.ListIndex = 0
   End If
   
End Sub

Public Sub FillCbAreaNeg(Cb As Object, ByVal Vigente As Boolean, Optional ByVal ItemBlanco As Boolean = True, Optional ByVal MaxItems As Integer = 0, Optional ByVal MsgMaxItem As String = "")
   Dim Rs As Recordset
   Dim i As Integer
   Dim CondVigente As String
   
   Cb.Clear
   
   If ItemBlanco Then
      Call AddItem(Cb, " ", 0)
   End If
   
   'Cb.AddItem " "
   'Cb.ItemData(Cb.NewIndex) = 0
   If Vigente Then
      CondVigente = " AND Vigente <> 0 "
   End If
         
   Set Rs = OpenRs(DbMain, "SELECT IdAreaNegocio, Descripcion FROM AreaNegocio WHERE IdEmpresa = " & gEmpresa.id & CondVigente)
   i = 1
   
   Do While Rs.EOF = False
   
      If MaxItems > 0 And i > MaxItems Then
         If MsgMaxItem <> "" Then
            MsgBox1 "Atención: " & MsgMaxItem, vbExclamation + vbOKOnly
            Exit Do
         End If
      End If
      
      Call AddItem(Cb, vFld(Rs("Descripcion"), True), vFld(Rs("IdAreaNegocio")))
      'Cb.AddItem vFld(Rs("Descripcion"), True)
      'Cb.ItemData(Cb.NewIndex) = vFld(Rs("IdAreaNegocio"))
      Rs.MoveNext
      
      i = i + 1
   Loop
   
   Call CloseRs(Rs)
End Sub
Public Sub FillCbCCosto(Cb As Object, ByVal Vigente As Boolean, Optional ByVal ItemBlanco As Boolean = True, Optional ByVal MaxItems As Integer = 0, Optional ByVal MsgMaxItem As String = "")
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim CondVigente As String
   
   Cb.Clear
   
   If Vigente Then
      CondVigente = " AND Vigente <> 0 "
   End If
   
   Cb.Clear
   
   If ItemBlanco Then
      Call AddItem(Cb, " ", 0)
   End If
   
   'Cb.AddItem " "
   'Cb.ItemData(Cb.NewIndex) = 0
   
   Q1 = "SELECT IdCCosto, Descripcion FROM CentroCosto"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & CondVigente
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   
   Do While Rs.EOF = False
   
      If MaxItems > 0 And i > MaxItems Then
         If MsgMaxItem <> "" Then
            MsgBox1 "Atención: " & MsgMaxItem, vbExclamation + vbOKOnly
            Exit Do
         End If
      End If
      
      Call AddItem(Cb, vFld(Rs("Descripcion"), True), vFld(Rs("IdCCosto")))
      'Cb.AddItem vFld(Rs("Descripcion"), True)
      'Cb.ItemData(Cb.NewIndex) = vFld(Rs("IdCCosto"))
      Rs.MoveNext
      
      i = i + 1
   Loop
   
   Call CloseRs(Rs)
   
End Sub
'PS
Public Function ChkOpt(ByVal Opciones As Long, ByVal Opt As Long)

   ChkOpt = ((Opciones And Opt) <> 0)

End Function
Public Sub AppendLogImpreso(ByVal IdLibOf As Long, Optional ByVal Mes As Integer = 0, Optional ByVal FDesde As Long = 0, Optional ByVal FHasta As Long = 0)
   Dim Q1 As String
   
   Q1 = "INSERT INTO LogImpreso (Fecha, IdEmpresa, Ano, IdInforme, IdUsuario, Mes, FDesde, FHasta, Estado ) "
   Q1 = Q1 & "VALUES(" & CLng(Int(Now))
   Q1 = Q1 & "," & gEmpresa.id
   Q1 = Q1 & "," & gEmpresa.Ano
   Q1 = Q1 & "," & IdLibOf
   Q1 = Q1 & "," & gUsuario.IdUsuario
   Q1 = Q1 & "," & Mes
   Q1 = Q1 & "," & FDesde
   Q1 = Q1 & "," & FHasta
   Q1 = Q1 & "," & EL_IMPRESO & ")"
   
   Call ExecSQL(DbMain, Q1)
   
End Sub

Public Function QryLogImpreso(ByVal IdLibOf As Long, ByVal Mes As Integer, FDesde As Long, FHasta As Long, Fecha As Long, Usuario As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   QryLogImpreso = False
   
   Q1 = "SELECT Fecha, Usuarios.NombreLargo, FDesde, FHasta "
   Q1 = Q1 & " FROM LogImpreso INNER JOIN Usuarios ON LogImpreso.IdUsuario = Usuarios.idUsuario "
   Q1 = Q1 & " WHERE IdInforme = " & IdLibOf & " AND Estado <> " & EL_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   If Mes > 0 Then
      Q1 = Q1 & " AND Mes = " & Mes
   End If
   
   Q1 = Q1 & " ORDER BY Fecha DESC"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      QryLogImpreso = True
      
      FDesde = vFld(Rs("FDesde"))
      FHasta = vFld(Rs("FHasta"))
      Fecha = vFld(Rs("Fecha"))
      Usuario = vFld(Rs("NombreLargo"), True)
      
   End If
   
   Call CloseRs(Rs)

End Function
Public Function LibMensualesImpresos(ByVal Mes As Integer, Optional ByVal Msg As Boolean = False) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim LibNoImpresos As String
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   
   'vemos si están impresos todos los libros oficiales mensuales
   If QryLogImpreso(LIBOF_COMPRAS, Mes, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_COMPRAS) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_VENTAS, Mes, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_VENTAS) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_RETEN, Mes, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_RETEN) & vbNewLine
   End If
   
   LibMensualesImpresos = (LibNoImpresos = "")
   
   If LibNoImpresos <> "" And Msg = True Then
      If MsgBox1("Los siguientes Libros Oficiales de este mes no han sido impresos:" & vbNewLine & vbNewLine & LibNoImpresos & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         LibMensualesImpresos = False
      Else
         LibMensualesImpresos = True
      End If
   End If

End Function
Public Function LibAnualesImpresos(Optional ByVal Msg As Boolean = False) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim LibNoImpresos As String
   Dim NLibCompras As Integer
   Dim NLibVentas As Integer
   Dim NLibReten As Integer
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   
   'vemos si están impresos los libros mensuales para todos los meses
   
   'Compras
   Q1 = "SELECT DISTINCT Mes FROM LogImpreso "
   Q1 = Q1 & " WHERE IdInforme = " & LIBOF_COMPRAS & " AND Estado <> " & EL_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   NLibCompras = 0
   Do While Rs.EOF = False
      NLibCompras = NLibCompras + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If NLibCompras < 12 Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_COMPRAS) & vbNewLine
   End If
   
   'Ventas
   Q1 = "SELECT DISTINCT Mes FROM LogImpreso "
   Q1 = Q1 & " WHERE IdInforme = " & LIBOF_VENTAS & " AND Estado <> " & EL_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   NLibVentas = 0
   Do While Rs.EOF = False
      NLibVentas = NLibVentas + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If NLibVentas < 12 Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_VENTAS) & vbNewLine
   End If
   
   'Retenciones
   Q1 = "SELECT DISTINCT Mes FROM LogImpreso "
   Q1 = Q1 & " WHERE IdInforme = " & LIBOF_RETEN & " AND Estado <> " & EL_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   NLibReten = 0
   Do While Rs.EOF = False
      NLibReten = NLibReten + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If NLibReten < 12 Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_RETEN) & vbNewLine
   End If
      
   'vemos si están impresos todos los libros oficiales anuales
   If QryLogImpreso(LIBOF_DIARIO, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_DIARIO) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_MAYOR, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_MAYOR) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_INVBAL, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_INVBAL) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_TRIBUTARIO, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_TRIBUTARIO) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_CLASIFICADO, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_CLASIFICADO) & vbNewLine
   End If
   
   If QryLogImpreso(LIBOF_COMPYSALDOS, 0, FDesde, FHasta, Fecha, Usuario) = False Then
      LibNoImpresos = LibNoImpresos & gLibroOficial(LIBOF_COMPYSALDOS) & vbNewLine
   End If
   
   LibAnualesImpresos = True
   
   If LibNoImpresos <> "" And Msg = True Then
      If MsgBox1("Los siguientes Libros Oficiales no han sido impresos:" & vbNewLine & vbNewLine & LibNoImpresos & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         LibAnualesImpresos = False
      End If
   End If

End Function

'genera saldos de apertura de cuentas a partir de año anterior
Public Function GenSaldosApertura(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim ConnStr As String
   Dim TmpTbl As String
   Dim TblCuentas As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
   RutMdb = Rut & ".mdb"
   
   GenSaldosApertura = False
   
   If gEmpresa.TieneAnoAntAccess Then  'los saldos ya fueron generados al crear el nuevo año desde Access
      GenSaldosApertura = True
      Exit Function
   End If
   
   If gEmprSeparadas Then
      If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
         If Msg Then
            MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar saldos de apertura en forma automática.", vbExclamation + vbOKOnly
         End If
         Exit Function
      End If
   End If
   
   'vemos si el año anterior está cerrado
   FCierre = -1
   
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar saldos de apertura.", vbExclamation + vbOKOnly
      End If
      Exit Function
   ElseIf FCierre < 0 Then
   
      If Msg Then
         MsgBox1 "El año anterior no existe. No es posible generar los saldos de apertura en forma automática.", vbExclamation + vbOKOnly
      End If
      Exit Function
   
   End If
   
#If DATACON = 1 Then       'Access
   
   If gEmprSeparadas Then
      'cerramos el año actual y abrimos el año anterior
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano - 1)
      Call LinkMdbAdm
   
   
      'linkeamos la tabla de cuentas del año actual para actualizar los saldos de apertura, a partir del año anterior
      'ConnStr = "PWD=" & PASSW_PREFIX & Rut & ";"
      'Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , ConnStr)
      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , gEmpresa.ConnStr)
   End If
   
#End If

   TmpTbl = DbGenTmpName2(gDbType, "tapertura_")
   
   'Generamos los saldos de apertura financiera
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
   'generamos tabla temporal con saldos
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Call ExecSQL(DbMain, Q1)
   
   'limpiamos los saldos de apertura
   If gEmprSeparadas Then
      TblCuentas = "CuentasNew"
   Else
      TblCuentas = "Cuentas"
   End If
   
   Q1 = "UPDATE " & TblCuentas & " SET Debe = 0, Haber = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los saldos con tabla temporal
'   Q1 = "UPDATE " & TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo "
'   Q1 = Q1 & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano "
'   Q1 = Q1 & "  SET " & TblCuentas & ".Debe = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
'   Q1 = Q1 & ", " & TblCuentas & ".Haber = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
'   Q1 = Q1 & " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   Tbl = TblCuentas
   sFrom = TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo "
   sFrom = sFrom & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano "
   sSet = TblCuentas & ".Debe = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
   sSet = sSet & ", " & TblCuentas & ".Haber = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
   sWhere = " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Generamos los saldos de apertura tributaria
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
   'generamos tabla temporal con saldos tributarios
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Call ExecSQL(DbMain, Q1)
   
   'limpiamos los saldos de apertura tributaria
   Q1 = "UPDATE " & TblCuentas & " SET DebeTrib = 0, HaberTrib = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los saldos de apertura tributaria con tabla temporal
'   Q1 = "UPDATE " & TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo"
'   Q1 = Q1 & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano"
'   Q1 = Q1 & "  SET " & TblCuentas & ".DebeTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
'   Q1 = Q1 & ", " & TblCuentas & ".HaberTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
'   Q1 = Q1 & " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   Tbl = TblCuentas
   sFrom = TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo"
   sFrom = sFrom & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano"
   sSet = TblCuentas & ".DebeTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
   sSet = sSet & ", " & TblCuentas & ".HaberTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
   sWhere = " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
   
   If gEmprSeparadas Then
      Q1 = "DROP TABLE CuentasNew"
      Call ExecSQL(DbMain, Q1)
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
   End If
   
   GenSaldosApertura = True

End Function

'genera saldos de apertura de cuentas a partir de año anterior
Public Function GenSaldosAperturaFull(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim Q2 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim ConnStr As String
   Dim TmpTbl As String
   Dim TmpTbl2 As String
   Dim TmpTbl3 As String
   Dim TblCuentas As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
   RutMdb = Rut & ".mdb"
   
   GenSaldosAperturaFull = False
   
   If gEmpresa.TieneAnoAntAccess Then  'los saldos ya fueron generados al crear el nuevo año desde Access
      GenSaldosAperturaFull = True
      Exit Function
   End If
   
   If gEmprSeparadas Then
      If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
         If Msg Then
            MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar saldos de apertura en forma automática.", vbExclamation + vbOKOnly
         End If
         Exit Function
      End If
   End If
   
   'vemos si el año anterior está cerrado
   FCierre = -1
   
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar saldos de apertura.", vbExclamation + vbOKOnly
      End If
      Exit Function
   ElseIf FCierre < 0 Then
   
      If Msg Then
         MsgBox1 "El año anterior no existe. No es posible generar los saldos de apertura en forma automática.", vbExclamation + vbOKOnly
      End If
      Exit Function
   
   End If
   
#If DATACON = 1 Then       'Access
   
   If gEmprSeparadas Then
      'cerramos el año actual y abrimos el año anterior
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano - 1)
      Call LinkMdbAdm
   
   
      'linkeamos la tabla de cuentas del año actual para actualizar los saldos de apertura, a partir del año anterior
      'ConnStr = "PWD=" & PASSW_PREFIX & Rut & ";"
      'Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , ConnStr)
      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , gEmpresa.ConnStr)
   End If
   
#End If

   TmpTbl = DbGenTmpName2(gDbType, "tapertura_")
   
   'Generamos los saldos de apertura financiera
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
'   'feña
   TmpTbl2 = DbGenTmpName2(gDbType, "tapertura_2")
   Q1 = "DROP TABLE " & TmpTbl2
   Call ExecSQL(DbMain, Q1)
'   'fin feña
   
   'generamos tabla temporal con saldos
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl2
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp = Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Call ExecSQL(DbMain, Q1)
   
   'feña
   Q2 = Q1
   TmpTbl3 = DbGenTmpName2(gDbType, "tapertura_3")
   Q1 = "DROP TABLE " & TmpTbl3
   Call ExecSQL(DbMain, Q1)
   
   Q2 = Replace(Replace(Replace(Replace(Q2, " Comprobante ", " ComprobanteFull "), " Comprobante.", " ComprobanteFull."), "MovComprobante", "MovComprobanteFull"), TmpTbl2, TmpTbl3)
   Call ExecSQL(DbMain, Q2)
   
   Q1 = " SELECT TM1.IDCUENTA as IdCuenta, TM1.CODIGO as Codigo, TM1.IDEMPRESA as IdEmpresa, TM1.ANO as Ano, (TM1.SUMDEBE + IIF(TM2.SUMDEBE IS NULL,0,TM2.SUMDEBE)) as SumDebe, (TM1.SUMHABER + IIF(TM2.SUMHABER IS NULL,0,TM2.SUMHABER)) as SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM " & TmpTbl2 & " AS TM1 LEFT JOIN " & TmpTbl3 & " AS TM2 ON TM1.IDCUENTA = TM2.IDCUENTA"
   Call ExecSQL(DbMain, Q1)
   'fin feña

   'limpiamos los saldos de apertura
   If gEmprSeparadas Then
      TblCuentas = "CuentasNew"
   Else
      TblCuentas = "Cuentas"
   End If
   
   Q1 = "UPDATE " & TblCuentas & " SET Debe = 0, Haber = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los saldos con tabla temporal
'   Q1 = "UPDATE " & TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo "
'   Q1 = Q1 & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano "
'   Q1 = Q1 & "  SET " & TblCuentas & ".Debe = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
'   Q1 = Q1 & ", " & TblCuentas & ".Haber = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
'   Q1 = Q1 & " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   Tbl = TblCuentas
   sFrom = TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo "
   sFrom = sFrom & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano "
   sSet = TblCuentas & ".Debe = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
   sSet = sSet & ", " & TblCuentas & ".Haber = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
   sWhere = " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Generamos los saldos de apertura tributaria
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
   'feña
   Q1 = "DROP TABLE " & TmpTbl2
   Call ExecSQL(DbMain, Q1)
   Q1 = "DROP TABLE " & TmpTbl3
   Call ExecSQL(DbMain, Q1)
   'fin feña
   'generamos tabla temporal con saldos tributarios
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl2
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp = Comprobante.idComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   
   Call ExecSQL(DbMain, Q1)
   
   'feña
   Q2 = Q1
   Q2 = Replace(Replace(Replace(Replace(Q2, " Comprobante ", " ComprobanteFull "), " Comprobante.", " ComprobanteFull."), "MovComprobante", "MovComprobanteFull"), TmpTbl2, TmpTbl3)
   Call ExecSQL(DbMain, Q2)
   
   Q1 = " SELECT TM1.IDCUENTA as IdCuenta, TM1.CODIGO as Codigo, TM1.IDEMPRESA as IdEmpresa, TM1.ANO as Ano, (TM1.SUMDEBE + IIF(TM2.SUMDEBE IS NULL,0,TM2.SUMDEBE)) as SumDebe, (TM1.SUMHABER + IIF(TM2.SUMHABER IS NULL,0,TM2.SUMHABER)) as SumHaber "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM " & TmpTbl2 & " AS TM1 LEFT JOIN " & TmpTbl3 & " AS TM2 ON TM1.IDCUENTA = TM2.IDCUENTA"
   Call ExecSQL(DbMain, Q1)
   'fin feña
   
   'limpiamos los saldos de apertura tributaria
   Q1 = "UPDATE " & TblCuentas & " SET DebeTrib = 0, HaberTrib = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los saldos de apertura tributaria con tabla temporal
'   Q1 = "UPDATE " & TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo"
'   Q1 = Q1 & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano"
'   Q1 = Q1 & "  SET " & TblCuentas & ".DebeTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
'   Q1 = Q1 & ", " & TblCuentas & ".HaberTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
'   Q1 = Q1 & " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   Tbl = TblCuentas
   sFrom = TblCuentas & " INNER JOIN " & TmpTbl & " ON " & TblCuentas & ".Codigo = " & TmpTbl & ".Codigo"
   sFrom = sFrom & " AND " & TblCuentas & ".IdEmpresa = " & TmpTbl & ".IdEmpresa AND " & TblCuentas & ".Ano - 1 = " & TmpTbl & ".Ano"
   sSet = TblCuentas & ".DebeTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, " & TmpTbl & ".SumDebe - " & TmpTbl & ".SumHaber,0) "
   sSet = sSet & ", " & TblCuentas & ".HaberTrib = iif( " & TmpTbl & ".SumDebe > " & TmpTbl & ".SumHaber, 0, " & TmpTbl & ".SumHaber - " & TmpTbl & ".SumDebe)"
   sWhere = " WHERE " & TblCuentas & ".IdEmpresa = " & IdEmpresa & " AND " & TblCuentas & ".Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   Q1 = "DROP TABLE " & TmpTbl
   Call ExecSQL(DbMain, Q1)
   
   'feña
   Q1 = "DROP TABLE " & TmpTbl2
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "DROP TABLE " & TmpTbl3
   Call ExecSQL(DbMain, Q1)
   'fin feña
   
   If gEmprSeparadas Then
      Q1 = "DROP TABLE CuentasNew"
      Call ExecSQL(DbMain, Q1)
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
   End If
   
   GenSaldosAperturaFull = True

End Function
Public Function GenCompApertura(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal GenSaldosAp As Boolean = True) As Boolean
   Dim NumCompAper As Long
   Dim IdCuentaResul As Long
   Dim IdCompAper As Long, IdCompAperTrib As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim Saldo As Double
   Dim Debe As Double
   Dim Haber As Double
   Dim Frm As FrmApertura
   Dim Rc As Integer
   Dim ResDebe As Double
   Dim ResHaber As Double
   Dim SaldoRes As Double
   Dim NAper As Long
   
   GenCompApertura = False
   
   'pedimos al usuario el nº de comp. de apertura, si corresponde, y la cuenta de resultado
   Set Frm = New FrmApertura
   Rc = Frm.FSelect(IdEmpresa, Ano, NumCompAper, IdCompAper, IdCuentaResul, IdCompAperTrib)
   Set Frm = Nothing
   
   If Rc = vbCancel Then   'se arrepintió de generar el comprobante de apertura
   
      'vemos si no hay comprobante de apertura ya generado
   
      Q1 = "SELECT IdCompAper "
      Q1 = Q1 & " FROM EmpresasAno "
      Q1 = Q1 & " WHERE idEmpresa=" & IdEmpresa & " AND Ano=" & Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         IdCompAper = vFld(Rs("IdCompAper"))
      End If
      
      Call CloseRs(Rs)

      If IdCompAper = 0 Then  'se supone que no hay uno generado (y por lo tanto no hay un comprobante de apertura generado tampoco)
         
         'verifiquemos por si acaso si hay algún comprobante de tipo Apertura.
         'Si es así, se elimina (esto no debiera ocurrir nunca)
         Call VerificaMultiCompApertura
      
         'ahora no hay uno generado, generamos uno en blanco de cada tipo: financiero y tributario, para guardar el número
         Call GenCompAperSinMovs(1, IdEmpresa, Ano, IdCompAperTrib)
         
      End If
   
      Exit Function
   End If
         
   If NumCompAper = 0 Then
      NumCompAper = 1
   End If
   
   'generamos saldos de apertura, tanto financiero como tributario

   If GenSaldosAp Then
      If GenSaldosApertura(IdEmpresa, Rut, Ano, True) = False Then
         Exit Function
      Else
         MsgBox1 "Se calcularon los saldos de apertura.", vbInformation
      End If
   End If
   
   '***------ generamos comprobante de apertura financiero  -------***
   
   
   'vemos si hay diferencia en los saldos de apertura financieros
   Q1 = "SELECT Sum(Debe) as SumDebe, Sum(Haber) as SumHaber "
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("SumDebe"))
      TotHaber = vFld(Rs("SumHaber"))
   End If
   
   Call CloseRs(Rs)
   
   Saldo = TotDebe - TotHaber
   
   If Saldo <> 0 Then
   
      'hay diferencia, ajustamos saldo de apertura de la cuenta de resultado
      
      Q1 = "SELECT Debe, Haber FROM Cuentas "
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      ResDebe = 0
      ResHaber = 0
      SaldoRes = 0
      
      If Rs.EOF = False Then
         ResDebe = vFld(Rs("Debe"))
         ResHaber = vFld(Rs("Haber"))
      End If
      
      Call CloseRs(Rs)
      
      SaldoRes = (ResDebe - ResHaber) - Saldo
      
      Q1 = "UPDATE Cuentas SET "
      
      If SaldoRes > 0 Then
         Q1 = Q1 & "  Debe = " & SaldoRes
         Q1 = Q1 & ", Haber = 0"
      Else
         Q1 = Q1 & "  Debe = 0"
         Q1 = Q1 & ", Haber = " & Abs(SaldoRes)
      End If
      
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      
      Call ExecSQL(DbMain, Q1)
   End If
   
   'con los saldos de apertura iguales, generamos comprobante de apertura financiero
   If IdCompAper > 0 And IdCompAperTrib > 0 Then
      Q1 = "  WHERE IdComp = " & IdCompAper
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "MovComprobante", Q1)
      
      Q1 = " WHERE IdComp = " & IdCompAperTrib
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "MovComprobante", Q1)
   
   Else           'nuevo comprobante, lo agregamos
   
      'antes verifiquemos por si acaso si hay algún comprobante de tipo Apertura (financiero o tributario).
      'Si es así, se elimina (esto no debiera ocurrir nunca)
      Call VerificaMultiCompApertura

      IdCompAper = GenCompAperSinMovs(NumCompAper, IdEmpresa, Ano, IdCompAperTrib)
      
   End If

   'insertamos movs comprobante, de acuerdo a los saldos de apertura
   Q1 = "INSERT INTO MovComprobante (IdComp, Orden, IdCuenta, Debe, Haber, Glosa, IdEmpresa, Ano)"
   Q1 = Q1 & "  SELECT " & IdCompAper & " As IdComp "
   Q1 = Q1 & ", 1 as Orden, IdCuenta As IdCuenta "
   Q1 = Q1 & ", Debe as Debe, Haber as Haber "
   Q1 = Q1 & ", 'Apertura' as Glosa, " & gEmpresa.id & " As IdEmpresa," & Ano & " As Ano"
   Q1 = Q1 & "  FROM Cuentas WHERE (Debe-Haber) <> 0 "
   Q1 = Q1 & "  AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & "  ORDER BY Cuentas.Codigo "
   
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos totales comprobante de apertura
   Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante "
   Q1 = Q1 & " WHERE IdComp = " & IdCompAper
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano

   Set Rs = OpenRs(DbMain, Q1)

   TotDebe = 0
   TotHaber = 0
   Saldo = 0
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
      Saldo = TotDebe - TotHaber
   End If

   Call CloseRs(Rs)

   'Resultado(Ajuste) para el comprobante de apertura
   'Actualizamos el TotalDebe y TotalHaber del Comprobante de Apertura
   Q1 = "UPDATE Comprobante SET "
   Q1 = Q1 & " Estado=" & EC_APROBADO     'por si lo hubieran anulado
   Q1 = Q1 & ", TipoAjuste=" & TAJUSTE_FINANCIERO
   Q1 = Q1 & ",TotalDebe = " & TotDebe
   Q1 = Q1 & ",TotalHaber=" & TotHaber
   Q1 = Q1 & " WHERE idComp=" & IdCompAper
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)

   If Saldo <> 0 Then      'esto no debiera ocurrir nunca, ya que ya se hizo el ajuste en la tabla de cuentas
      MsgBox1 "Error al generar comprobante de apertura. Saldo de Debe y Haber no son iguales.", vbExclamation + vbOKOnly
   End If

   Call AddLogComprobantes(IdCompAper, gUsuario.IdUsuario, O_EDIT, Now, EC_APROBADO, NumCompAper, CLng(DateSerial(Ano, 1, 1)), TC_APERTURA, EC_APROBADO, TAJUSTE_FINANCIERO)

   
   '***------ generamos comprobante de apertura trubutario  -------***
   
   
   'vemos si hay diferencia en los saldos de apertura tributario
   Q1 = "SELECT Sum(DebeTrib) as SumDebeTrib, Sum(HaberTrib) as SumHaberTrib "
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("SumDebeTrib"))
      TotHaber = vFld(Rs("SumHaberTrib"))
   End If
   
   Call CloseRs(Rs)
   
   Saldo = TotDebe - TotHaber
   
   If Saldo <> 0 Then
   
      'hay diferencia, ajustamos saldo de apertura de la cuenta de resultado
      
      Q1 = "SELECT DebeTrib, HaberTrib FROM Cuentas "
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      ResDebe = 0
      ResHaber = 0
      SaldoRes = 0
      
      If Rs.EOF = False Then
         ResDebe = vFld(Rs("DebeTrib"))
         ResHaber = vFld(Rs("HaberTrib"))
      End If
      
      Call CloseRs(Rs)
      
      SaldoRes = (ResDebe - ResHaber) - Saldo
      
      Q1 = "UPDATE Cuentas SET "
      
      If SaldoRes > 0 Then
         Q1 = Q1 & "  DebeTrib = " & SaldoRes
         Q1 = Q1 & ", HaberTrib = 0"
      Else
         Q1 = Q1 & "  DebeTrib = 0"
         Q1 = Q1 & ", HaberTrib = " & Abs(SaldoRes)
      End If
      
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul     'se utiliza la misma cuenta para el comprobante financiero y el tributario
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Call ExecSQL(DbMain, Q1)
   End If
   

   'insertamos movs comprobante, de acuerdo a los saldos de apertura
   Q1 = "INSERT INTO MovComprobante (IdComp, Orden, IdCuenta, Debe, Haber, Glosa, IdEmpresa, Ano)"
   Q1 = Q1 & "  SELECT " & IdCompAperTrib & " As IdComp "
   Q1 = Q1 & ", 1 as Orden, IdCuenta As IdCuenta "
   Q1 = Q1 & ", DebeTrib as Debe, HaberTrib as Haber "
   Q1 = Q1 & ", 'Apertura' as Glosa, " & gEmpresa.id & " As IdEmpresa," & Ano & " As Ano"
   Q1 = Q1 & "  FROM Cuentas WHERE (DebeTrib-HaberTrib) <> 0 "
   Q1 = Q1 & "  AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & "  ORDER BY Cuentas.Codigo "
   
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos totales comprobante de apertura
   Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante "
   Q1 = Q1 & " WHERE IdComp = " & IdCompAperTrib
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano

   Set Rs = OpenRs(DbMain, Q1)

   TotDebe = 0
   TotHaber = 0
   Saldo = 0
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
      Saldo = TotDebe - TotHaber
   End If

   Call CloseRs(Rs)

   'Resultado(Ajuste) para el comprobante de apertura
   'Actualizamos el TotalDebe y TotalHaber del Comprobante de Apertura
   Q1 = "UPDATE Comprobante SET "
   Q1 = Q1 & " Estado=" & EC_APROBADO     'por si lo hubieran anulado
   Q1 = Q1 & ",TipoAjuste=" & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & ",TotalDebe = " & TotDebe
   Q1 = Q1 & ",TotalHaber=" & TotHaber
   Q1 = Q1 & " WHERE idComp=" & IdCompAperTrib
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)

   If Saldo <> 0 Then      'esto no debiera ocurrir nunca, ya que ya se hizo el ajuste en la tabla de cuentas
      MsgBox1 "Error al generar comprobante de apertura. Saldo de Debe y Haber no son iguales.", vbExclamation + vbOKOnly
   End If

   Call AddLogComprobantes(IdCompAper, gUsuario.IdUsuario, O_EDIT, Now, EC_APROBADO, NumCompAper, CLng(DateSerial(Ano, 1, 1)), TC_APERTURA, EC_APROBADO, TAJUSTE_TRIBUTARIO)
   
   GenCompApertura = True
   
End Function
Public Function RecalcSaldos(ByVal IdEmpresa As Long, ByVal Ano As Integer, Optional ByVal bIniNull As Boolean = 1)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   'PIPE otros doc
   Dim Rs2 As Recordset
   Dim Debe As Double
   Dim Haber As Double
   Dim Saldo As Double
   Dim CurIdDoc As Long
   Dim WhLib As String
   
   '2931541
   Dim PagoAnoAnterior As Double
   '2931541
   
   '3025162
   Dim RsCopy As Object
   '3025162
   
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim DocOtroEsCargo As Boolean
   Dim SetPagado As Boolean
   
   If IdEmpresa = 0 Then
      IdEmpresa = gEmpresa.id
   End If
   
   If Ano = 0 Then
      Ano = gEmpresa.Ano
   End If
   
   '3125609
   Call limpiarCampoTotPagado
   '3125609
   
   
   WhLib = " Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") "
   
   'marcamos notas de crédito y débito con SaldoDoc = NULL que tienen factura asociada con SaldoDoc = NULL
'   Q1 = "UPDATE Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
'   Q1 = Q1 & " SET Documento.SaldoDoc = NULL WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   sFrom = " Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Documento_1", True)  ' " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
   sSet = " Documento.SaldoDoc = NULL "
   sWhere = " WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   If bIniNull Then ' 14 feb 2020: ya se le puso NULL justo antes de llamar
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If
               
               
   'consulta de docs que no están enlazados a ningún comprobante
   Q1 = " SELECT 1, Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
   Q1 = Q1 & " Sum(MovDocumento.Debe) As Debe, Sum(MovDocumento.Haber) As Haber "
   
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541
   
   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609
   
   Q1 = Q1 & " FROM ((Documento  "
   Q1 = Q1 & "  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento", True) & " )" ' 14 feb 2020: se agrega , True
   'Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc"
   Q1 = Q1 & "  LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
'   Q1 = Q1 & "   AND Documento.IdEmpresa = vMovCompIdDoc.IdEmpresa AND Documento.Ano = vMovCompIdDoc.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc", True) & " )" ' 14 feb 2020
   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
'   Q1 = Q1 & "   AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano "
   
     '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541

   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1", True) ' 14 feb 2020
  
   Q1 = Q1 & " WHERE " & WhLib & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL) "
   'Q1 = Q1 & "  AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND Documento.Estado <> " & ED_ANULADO
         'tomamos los que no están enlazados a un comprobante y los que están marcados como centralizados pero no tienen comprobante asociado (docs pendientes del año anterior)
   'Q1 = Q1 & "  AND (MovComprobante.IdComp IS NULL "
   Q1 = Q1 & "  AND (vMovCompIdDoc.IdDoc IS NULL "
   Q1 = Q1 & "  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   
   Q1 = Q1 & " GROUP BY Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
   
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   
   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609
   
   Q1 = Q1 & " UNION "

   'consulta de movs. comprobantes que tienen docs enlazados
   Q1 = Q1 & " SELECT 2, Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, 0 as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
   Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) As Haber "

     '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541

   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609

   Q1 = Q1 & " FROM ((MovComprobante INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & "  INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
   '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1")

   Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO
   'Q1 = Q1 & " AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND " & WhLib
   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   
   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609

   Q1 = Q1 & " UNION "

   'consulta de docs ASOCIADOS que no están enlazados a ningún comprobante
   Q1 = Q1 & " SELECT 3, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0, "
   Q1 = Q1 & " Sum(MovDocumento.Debe) AS Debe, Sum(MovDocumento.Haber) AS Haber"
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541
   
   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609

   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
   '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc")


   Q1 = Q1 & " WHERE " & WhLib & " AND Documento.IdDocAsoc <> 0"
   Q1 = Q1 & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL)"
   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
   Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND (vMovCompIdDoc.IdDoc IS NULL  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   
   '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609
   
   Q1 = Q1 & " UNION "

   'consulta de movs. comprobantes que tienen docs ASOCIADOS enlazados

   Q1 = Q1 & " SELECT 4, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotPagadoAnoAnt, 0 AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0,  "
   Q1 = Q1 & " Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541
    
    '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609

   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
    '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
   Q1 = Q1 & " INNER JOIN MovComprobante ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")

   Q1 = Q1 & " WHERE Comprobante.Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND " & WhLib
   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
   
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   
    '3125609
   Q1 = Q1 & ", Documento.FExported "
   '3125609

   Q1 = Q1 & " ORDER BY IdDoc"

   Set Rs = OpenRs(DbMain, Q1)
      
  Do While Rs.EOF = False
        
      'detalle doc
      If CurIdDoc <> vFld(Rs("IdDoc")) Then   ' puede venir más de una vez cuando hay documentos asociados
      
        ' If vFld(Rs("IdDoc")) = 483 Then   'OJO CON EL ESTADO DEL DOCUEMENTO DEL AñO ANTERIOR QUE DEBE ESTAR CENTRALIZADO O PAGADO
'        If vFld(Rs("IdDoc")) = 19587 Then
'           MsgBeep vbExclamation
'
'        End If

         Debe = 0
         Haber = 0
         SetPagado = False
         
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
         
            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
               Saldo = vFld(Rs("Total"))
            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = (Debe - Haber)
                        
            End If
            
            '3047309
'            If vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
'               If Abs(vFld(Rs("Total"))) = Abs(Saldo) Then
'                 Saldo = Saldo + Abs(vFld(Rs("TotPagadoAnoAnt"))) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'               Else
'
'                Saldo = Abs(vFld(Rs("Total")) - (Saldo + vFld(Rs("TotPagadoAnoAnt")))) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'               End If
'        Else
                        
                Saldo = Saldo - (vFld(Rs("TotPagadoAnoAnt"))) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
                '2931541
                PagoAnoAnterior = vFld(Rs("TotPagadoAnoAnt"))
                '2931541
                        
                    
            
            'Saldo = Saldo - (IIf(vFld(Rs("TotPagadoAnoAnt")) > 0, vFld(Rs("TotPagadoAnoAnt")) * -1, vFld(Rs("TotPagadoAnoAnt"))))
        'End If
           '  3047309

'            Saldo = Saldo = Saldo - Abs(vFld(Rs("TotPagadoAnoAnt")))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            
            
            
         Else
         
         
'            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
'               Saldo = vFld(Rs("Total"))
'            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = Debe - Haber
            
               If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                  If Saldo <= 0 Then
                     DocOtroEsCargo = True
                  Else
                     DocOtroEsCargo = False
                  End If
               End If
            
'            End If
           
           
            If IsNull(Rs("TotPagadoAnoAnt")) = False Then    'FCA 04 feb 2020: se asume que sólo el primer año el TotPagadoAnoAnt es NULL, por lo tanto, al segundo año se le agrega el total
               Saldo = Saldo + IIf(vFld(Rs("DocOtroEsCargo")), -1, 1) * vFld(Rs("Total"))
            End If
            
           
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0 "
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
           
            Call ExecSQL(DbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
                              
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado, pero si ponemos el DocOtroEsCargo
                           
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               Else   'Saldo = Total
               
                  If vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano Then   'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
                                 
               End If
               
            Else
            '633824
'            If vFld(Rs("TipoLib")) = LIB_REMU Then
'                If Saldo < 0 Then
'                     DocOtroEsCargo = True
'                  Else
'                     DocOtroEsCargo = False
'                  End If
'            End If
            '633824
               Q1 = Q1 & ", DocOtroEsCargo = " & Abs(DocOtroEsCargo)
                                      
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
           
            Call ExecSQL(DbMain, Q1)
         
         End If
         
      Else
                 
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
            
            '2931541
'            If (vFld(Rs("IdDoc")) = 299 And PagoAnoAnterior = 44458412 And vFld(Rs("IdEmpresa")) = 12) Or (vFld(Rs("IdDoc")) = 300 And PagoAnoAnterior = 38358588 And vFld(Rs("IdEmpresa")) = 12) Then
'            Debe = Debe + vFld(Rs("Debe")) - PagoAnoAnterior
'            Else
'            Debe = Debe + vFld(Rs("Debe"))
'            End If
            
            '2931541
            
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
           '3047309
'            If vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
'                Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt")) '- PagoAnoAnterior 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'            Else
'               Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt")) - PagoAnoAnterior 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'            End If
            '3047309
            
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))
            '2931541
         Else
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
            If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                If Saldo <= 0 Then
                   DocOtroEsCargo = True
                Else
                   DocOtroEsCargo = False
                End If
             End If
                                   
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0"
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
            Call ExecSQL(DbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
            
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado
               
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Or vFld(Rs("Estado")) = ED_PAGADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               ElseIf Not SetPagado Then    'Saldo = Total
                  
                  If (vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano) Then    'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
               
               End If
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            Call ExecSQL(DbMain, Q1)
         End If
         
      End If
         
      CurIdDoc = vFld(Rs("IdDoc"))
      ' 2828725 Cambia a estado pendiente los documentos NDV si los comprobantes no suman igual el debe y el haber
'        Q1 = "SELECT Switch(sum(debe) = sum(haber), 0,sum(debe) <> sum(haber),  abs(sum(debe) - sum(haber))) as pagado "
'        Q1 = Q1 & " FROM    documento as docu, movcomprobante as mov, comprobante com "
'        Q1 = Q1 & " WHERE   docu.iddoc = mov.iddoc "
'        Q1 = Q1 & " AND     mov.idcomp = com.idcomp "
'        Q1 = Q1 & " AND tipolib = 2 AND tipodoc = 4 "
'        Q1 = Q1 & " AND     docu.numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'        Set Rs2 = OpenRs(DbMain, Q1)
'
'
'        Do While Rs2.EOF = False
'
'            If vFld(Rs2("pagado")) > 0 Then
'
'            Q1 = "UPDATE documento "
'            Q1 = Q1 & " SET Estado = " & ED_PENDIENTE
'            Q1 = Q1 & " , SaldoDoc = " & vFld(Rs2("pagado"))
'            Q1 = Q1 & " WHERE numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'            Call ExecSQL(DbMain, Q1)
'
'            End If
'        Rs2.MoveNext
'        Loop
'        Call CloseRs(Rs2)
        ' FIN 2828725
               
      
      Rs.MoveNext
   Loop
      
   Call CloseRs(Rs)
   
   
   
End Function 'genera los docs que quedaron pendientes del año anterior
Public Function RecalcSaldosFulle(ByVal IdEmpresa As Long, ByVal Ano As Integer, Optional ByVal bIniNull As Boolean = 1)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   'PIPE otros doc
   Dim Rs2 As Recordset
   Dim Debe As Double
   Dim Haber As Double
   Dim Saldo As Double
   Dim CurIdDoc As Long
   Dim WhLib As String
  
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim DocOtroEsCargo As Boolean
   Dim SetPagado As Boolean
   
   If IdEmpresa = 0 Then
      IdEmpresa = gEmpresa.id
   End If
   
   If Ano = 0 Then
      Ano = gEmpresa.Ano
   End If
   
   
   WhLib = " Documento.TipoLib IN( " & LIB_OTROFULL & ") "
   
   'marcamos notas de crédito y débito con SaldoDoc = NULL que tienen factura asociada con SaldoDoc = NULL
'   Q1 = "UPDATE Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
'   Q1 = Q1 & " SET Documento.SaldoDoc = NULL WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   'sFrom = " DocumentoFull INNER JOIN DocumentoFull as DocumentoFull_1 ON DocumentoFull.IdDocAsoc = DocumentoFull_1.IdDoc "
   'sFrom = sFrom & JoinEmpAno(gDbType, "DocumentoFull", "DocumentoFull_1", True)  ' " AND DocumentoFull.IdEmpresa = DocumentoFull_1.IdEmpresa AND DocumentoFull.Ano = DocumentoFull_1.Ano"
   sFrom = " Documento "
   sSet = " Documento.SaldoDoc = NULL "
   sWhere = " WHERE " & WhLib '& " AND DocumentoFull.SaldoDoc IS NULL"
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   If bIniNull Then ' 14 feb 2020: ya se le puso NULL justo antes de llamar
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If
               
               
'   'consulta de docs que no están enlazados a ningún comprobante
'   Q1 = " SELECT 1, Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
'   'Q1 = Q1 & " Sum(MovDocumento.Debe) As Debe, Sum(MovDocumento.Haber) As Haber "
'   Q1 = Q1 & " iif(Documento.tratamiento = 1, Documento.Total, 0) AS Debe, iif(Documento.tratamiento = 2, Documento.Total, 0) AS Haber"
'
'
'   Q1 = Q1 & " FROM ((Documento  "
'   Q1 = Q1 & "  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento", True) & " )" ' 14 feb 2020: se agrega , True
'   'Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc"
'   Q1 = Q1 & "  LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
''   Q1 = Q1 & "   AND Documento.IdEmpresa = vMovCompIdDoc.IdEmpresa AND Documento.Ano = vMovCompIdDoc.Ano )"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc", True) & " )" ' 14 feb 2020
'   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
''   Q1 = Q1 & "   AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1", True) ' 14 feb 2020
'
'   Q1 = Q1 & " WHERE " & WhLib & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL) "
'   'Q1 = Q1 & "  AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND Documento.Estado <> " & ED_ANULADO
'         'tomamos los que no están enlazados a un comprobante y los que están marcados como centralizados pero no tienen comprobante asociado (docs pendientes del año anterior)
'   'Q1 = Q1 & "  AND (MovComprobante.IdComp IS NULL "
'   Q1 = Q1 & "  AND (vMovCompIdDoc.IdDoc IS NULL "
'   Q1 = Q1 & "  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, Documento.Tratamiento "
'
'   Q1 = Q1 & " UNION "
   
'   'consulta de movs. comprobantes que tienen docs enlazados
'   Q1 = Q1 & " SELECT 2, Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, 0 as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
'   'Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) As Haber "
'   Q1 = Q1 & " iif(Documento.tratamiento = 1, Documento.Total, 0) AS Debe, iif(Documento.tratamiento = 2, Documento.Total, 0) AS Haber"
'
'   Q1 = Q1 & " FROM ((MovComprobante INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
'   Q1 = Q1 & "  INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1")
'
'   Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO
'   'Q1 = Q1 & " AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND " & WhLib
'   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, Documento.Tratamiento "
'
'   Q1 = Q1 & " UNION "
   
'   'consulta de docs ASOCIADOS que no están enlazados a ningún comprobante
'   Q1 = Q1 & " SELECT 3, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0, "
'   'Q1 = Q1 & " Sum(MovDocumento.Debe) AS Debe, Sum(MovDocumento.Haber) AS Haber"
'   Q1 = Q1 & " iif(Documento.tratamiento = 1, Documento.Total, 0) AS Debe, iif(Documento.tratamiento = 2, Documento.Total, 0) AS Haber"
'
'   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
'   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
'   Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc")
'
'
'   Q1 = Q1 & " WHERE " & WhLib & " AND Documento.IdDocAsoc <> 0"
'   Q1 = Q1 & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL)"
'   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO
'   Q1 = Q1 & " AND (vMovCompIdDoc.IdDoc IS NULL  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'
'   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.Tratamiento "
'
'   Q1 = Q1 & " UNION "

   'consulta de movs. comprobantes que tienen docs ASOCIADOS enlazados
  
'   Q1 = Q1 & " SELECT 4, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotPagadoAnoAnt, 0 AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0,  "
'   'Q1 = Q1 & " Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
'   Q1 = Q1 & " iif(Documento.tratamiento = 1, Documento.Total, 0) AS Debe, iif(Documento.tratamiento = 2, Documento.Total, 0) AS Haber"
'
'   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
'   Q1 = Q1 & " INNER JOIN MovComprobante ON MovComprobante.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
'   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
'
'   Q1 = Q1 & " WHERE Comprobante.Estado <> " & ED_ANULADO
'   Q1 = Q1 & " AND " & WhLib
'   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.Tratamiento "
  
  
   Q1 = " SELECT 1, doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotPagadoAnoAnt, 0 AS MovIdDoc, doc.IdDocAsoc, doc.Estado As EstadoDocAsoc, doc.IdCompCent, doc.IdCompPago,  "
   'Q1 = Q1 & " Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   Q1 = Q1 & " doc.FEmision, 0,   iif(doc.tratamiento = 1, doc.Total, SUM(MovCom.debe)) AS Debe, iif(doc.tratamiento = 2, doc.Total, SUM(MovCom.Haber)) AS Haber "
   
   Q1 = Q1 & " FROM Documento doc LEFT JOIN MovComprobante MovCom ON doc.iddoc = MovCom.IdDoc  "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
'   Q1 = Q1 & " INNER JOIN MovComprobante ON MovComprobante.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
'   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   
   Q1 = Q1 & " WHERE doc.Estado <> " & ED_ANULADO
   'Q1 = Q1 & " AND " & WhLib
   Q1 = Q1 & " AND doc.SaldoDoc IS NULL"
   Q1 = Q1 & " AND doc.IdEmpresa = " & IdEmpresa & " AND doc.Ano = " & Ano
   
   Q1 = Q1 & " GROUP BY doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotpagadoAnoAnt, doc.IdDocAsoc, doc.Estado, doc.IdCompCent, doc.IdCompPago, doc.FEmision, doc.Tratamiento "
  

   Q1 = Q1 & " ORDER BY doc.IdDoc"

   Set Rs = OpenRs(DbMain, Q1)
      
      
   Do While Rs.EOF = False
                                
      
      'detalle doc
      If CurIdDoc <> vFld(Rs("IdDoc")) Then   ' puede venir más de una vez cuando hay documentos asociados
      
'         If vFld(Rs("IdDoc")) = 12871 Then   'OJO CON EL ESTADO DEL DOCUEMENTO DEL AñO ANTERIOR QUE DEBE ESTAR CENTRALIZADO O PAGADO
'            MsgBeep vbExclamation
'         End If

         Debe = 0
         Haber = 0
         SetPagado = False
         
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU And vFld(Rs("TipoLib")) <> LIB_OTROFULL Then
         
            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
               Saldo = vFld(Rs("Total"))
            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = (Debe - Haber)
                        
            End If
            
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         Else
         
         
'            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
'               Saldo = vFld(Rs("Total"))
'            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = Debe - Haber
            
               If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                  If Saldo <= 0 Then
                     DocOtroEsCargo = True
                  Else
                     DocOtroEsCargo = False
                  End If
               End If
            
'            End If
           
           
            If IsNull(Rs("TotPagadoAnoAnt")) = False And vFld(Rs("TipoLib")) <> LIB_OTROFULL Then     'FCA 04 feb 2020: se asume que sólo el primer año el TotPagadoAnoAnt es NULL, por lo tanto, al segundo año se le agrega el total
               Saldo = Saldo + IIf(vFld(Rs("DocOtroEsCargo")), -1, 1) * vFld(Rs("Total"))
            ElseIf IsNull(Rs("TotPagadoAnoAnt")) = False And vFld(Rs("TipoLib")) = LIB_OTROFULL Then
                Saldo = Saldo '+ IIf(vFld(Rs("DocOtroEsCargo")), -1, 1) * vFld(Rs("Total"))
            End If
            
           
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0 "
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
           
            Call ExecSQL(DbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
                              
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU And vFld(Rs("TipoLib")) <> LIB_OTROFULL Then    'para LIB_OTROS no cambiamos estado, pero si ponemos el DocOtroEsCargo
                           
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               Else   'Saldo = Total
               
                  If vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano Then   'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
                                 
               End If
               
            Else
               Q1 = Q1 & ", DocOtroEsCargo = " & Abs(DocOtroEsCargo)
                                      
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
            Call ExecSQL(DbMain, Q1)
         
         End If
         
      Else
                 
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU And vFld(Rs("TipoLib")) <> LIB_OTROFULL Then
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))  'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            
         Else
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
            If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                If Saldo <= 0 Then
                   DocOtroEsCargo = True
                Else
                   DocOtroEsCargo = False
                End If
             End If
                                   
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0"
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
            Call ExecSQL(DbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
            
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU And vFld(Rs("TipoLib")) <> LIB_OTROFULL Then    'para LIB_OTROS no cambiamos estado
               
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Or vFld(Rs("Estado")) = ED_PAGADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               ElseIf Not SetPagado Then    'Saldo = Total
                  
                  If (vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano) Then    'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
               
               End If
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            Call ExecSQL(DbMain, Q1)
         End If
         
      End If
         
      CurIdDoc = vFld(Rs("IdDoc"))
      ' 2828725 Cambia a estado pendiente los documentos NDV si los comprobantes no suman igual el debe y el haber
'        Q1 = "SELECT Switch(sum(debe) = sum(haber), 0,sum(debe) <> sum(haber),  abs(sum(debe) - sum(haber))) as pagado "
'        Q1 = Q1 & " FROM    documento as docu, movcomprobante as mov, comprobante com "
'        Q1 = Q1 & " WHERE   docu.iddoc = mov.iddoc "
'        Q1 = Q1 & " AND     mov.idcomp = com.idcomp "
'        Q1 = Q1 & " AND tipolib = 2 AND tipodoc = 4 "
'        Q1 = Q1 & " AND     docu.numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'        Set Rs2 = OpenRs(DbMain, Q1)
'
'
'        Do While Rs2.EOF = False
'
'            If vFld(Rs2("pagado")) > 0 Then
'
'            Q1 = "UPDATE documento "
'            Q1 = Q1 & " SET Estado = " & ED_PENDIENTE
'            Q1 = Q1 & " , SaldoDoc = " & vFld(Rs2("pagado"))
'            Q1 = Q1 & " WHERE numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'            Call ExecSQL(DbMain, Q1)
'
'            End If
'        Rs2.MoveNext
'        Loop
'        Call CloseRs(Rs2)
        ' FIN 2828725
               
      Rs.MoveNext
   Loop
      
   Call CloseRs(Rs)
   
End Function

Public Function RecalcSaldosFull(ByVal IdEmpresa As Long, ByVal Ano As Integer, Optional ByVal bIniNull As Boolean = 1)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   'PIPE otros doc
   Dim Rs2 As Recordset
   Dim Debe As Double
   Dim Haber As Double
   Dim Saldo As Double
   Dim CurIdDoc As Long
   Dim WhLib As String
  
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim DocOtroEsCargo As Boolean
   Dim SetPagado As Boolean
   
   If IdEmpresa = 0 Then
      IdEmpresa = gEmpresa.id
   End If
   
   If Ano = 0 Then
      Ano = gEmpresa.Ano
   End If
   
     
   
   WhLib = " Documento.TipoLib IN(" & LIB_OTROFULL & ") "
   
   'marcamos notas de crédito y débito con SaldoDoc = NULL que tienen factura asociada con SaldoDoc = NULL
'   Q1 = "UPDATE Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
'   Q1 = Q1 & " SET Documento.SaldoDoc = NULL WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   'sFrom = " DocumentoFull INNER JOIN DocumentoFull as DocumentoFull_1 ON DocumentoFull.IdDocAsoc = DocumentoFull_1.IdDoc "
   'sFrom = sFrom & JoinEmpAno(gDbType, "DocumentoFull", "DocumentoFull_1", True)  ' " AND DocumentoFull.IdEmpresa = DocumentoFull_1.IdEmpresa AND DocumentoFull.Ano = DocumentoFull_1.Ano"
   sFrom = " Documento "
   sSet = " Documento.SaldoDoc = NULL "
   sWhere = " WHERE " & WhLib '& " AND DocumentoFull.SaldoDoc IS NULL"
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   If bIniNull Then ' 14 feb 2020: ya se le puso NULL justo antes de llamar
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If
               
               
   'consulta de docs que no están enlazados a ningún comprobante
'   Q1 = " SELECT 1, Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
'   Q1 = Q1 & " Sum(MovDocumento.Debe) As Debe, Sum(MovDocumento.Haber) As Haber "
'
'
'
'   Q1 = Q1 & " FROM ((Documento  "
'   Q1 = Q1 & "  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento", True) & " )" ' 14 feb 2020: se agrega , True
'   'Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc"
'   Q1 = Q1 & "  LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
''   Q1 = Q1 & "   AND Documento.IdEmpresa = vMovCompIdDoc.IdEmpresa AND Documento.Ano = vMovCompIdDoc.Ano )"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc", True) & " )" ' 14 feb 2020
'   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
''   Q1 = Q1 & "   AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1", True) ' 14 feb 2020
'
'   Q1 = Q1 & " WHERE " & WhLib & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL) "
'   'Q1 = Q1 & "  AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND Documento.Estado <> " & ED_ANULADO
'         'tomamos los que no están enlazados a un comprobante y los que están marcados como centralizados pero no tienen comprobante asociado (docs pendientes del año anterior)
'   'Q1 = Q1 & "  AND (MovComprobante.IdComp IS NULL "
'   Q1 = Q1 & "  AND (vMovCompIdDoc.IdDoc IS NULL "
'   Q1 = Q1 & "  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
'
'   Q1 = Q1 & " UNION "
'
'   'consulta de movs. comprobantes que tienen docs enlazados
'   Q1 = Q1 & " SELECT 2, Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, 0 as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
'   Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) As Haber "
'
'   Q1 = Q1 & " FROM ((MovComprobante INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
'   Q1 = Q1 & "  INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
'   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1")
'
'   Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO
'   'Q1 = Q1 & " AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & "  AND " & WhLib
'   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
'   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
'
'   Q1 = Q1 & " UNION "
'
'   'consulta de docs ASOCIADOS que no están enlazados a ningún comprobante
'   Q1 = Q1 & " SELECT 3, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0, "
'   Q1 = Q1 & " Sum(MovDocumento.Debe) AS Debe, Sum(MovDocumento.Haber) AS Haber"
'
'   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
'   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
'   Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc")
'
'
'   Q1 = Q1 & " WHERE " & WhLib & " AND Documento.IdDocAsoc <> 0"
'   Q1 = Q1 & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL)"
'   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO
'   Q1 = Q1 & " AND (vMovCompIdDoc.IdDoc IS NULL  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'
'   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
'
'   Q1 = Q1 & " UNION "
'
'   'consulta de movs. comprobantes que tienen docs ASOCIADOS enlazados
'
'   Q1 = Q1 & " SELECT 4, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotPagadoAnoAnt, 0 AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0,  "
'   Q1 = Q1 & " Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
'
'   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
'   Q1 = Q1 & " INNER JOIN MovComprobante ON MovComprobante.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
'   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
'   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
'
'   Q1 = Q1 & " WHERE Comprobante.Estado <> " & ED_ANULADO
'   Q1 = Q1 & " AND " & WhLib
'   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'
'   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
'
'
'   Q1 = Q1 & " ORDER BY IdDoc"
    Q1 = " SELECT 1, doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotPagadoAnoAnt, 0 AS MovIdDoc, doc.IdDocAsoc, doc.Estado As EstadoDocAsoc, doc.IdCompCent, doc.IdCompPago, doc.FEmision, 0,   0 AS Debe, 0 AS Haber "
    Q1 = Q1 & " FROM Documento as doc "
    Q1 = Q1 & " Where doc.TipoLib IN(" & LIB_OTROFULL & ") "
    Q1 = Q1 & " AND doc.SaldoDoc IS NULL "
    Q1 = Q1 & " AND doc.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND doc.Ano = " & Ano
    Q1 = Q1 & " AND doc.TotPagadoAnoAnt > 0 "
    Q1 = Q1 & " GROUP BY doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotpagadoAnoAnt, doc.IdDocAsoc, doc.Estado, doc.IdCompCent, doc.IdCompPago, doc.FEmision "
    'Q1 = Q1 & " ORDER BY  doc.IdDoc "
    
    Q1 = Q1 & " UNION "
    
    
'    Q1 = Q1 & " SELECT 2, doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotPagadoAnoAnt, 0 AS MovIdDoc, doc.IdDocAsoc, doc.Estado As EstadoDocAsoc, doc.IdCompCent, doc.IdCompPago, doc.FEmision, 0,   Sum(mov.Debe) AS Debe, Sum(mov.Haber) AS Haber "
'    Q1 = Q1 & " FROM DocumentoFull as doc, ComprobanteFull as com, MovComprobanteFull as mov "
'    Q1 = Q1 & " Where doc.IdDoc = mov.IdDoc "
'    Q1 = Q1 & " AND com.idcomp = mov.idcomp "
'    Q1 = Q1 & " AND com.Estado <> " & ED_ANULADO
'    Q1 = Q1 & " AND  doc.TipoLib IN(" & LIB_OTROFULL & ") "
'    Q1 = Q1 & " AND doc.SaldoDoc IS NULL "
'    Q1 = Q1 & " AND doc.IdEmpresa = " & IdEmpresa
'    Q1 = Q1 & " AND doc.Ano = " & Ano
'    Q1 = Q1 & " GROUP BY doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotpagadoAnoAnt, doc.IdDocAsoc, doc.Estado, doc.IdCompCent, doc.IdCompPago, doc.FEmision "
'    Q1 = Q1 & " ORDER BY  doc.IdDoc "
    
    Q1 = Q1 & " SELECT 2, doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotPagadoAnoAnt, 0 AS MovIdDoc, doc.IdDocAsoc, doc.Estado As EstadoDocAsoc, doc.IdCompCent, doc.IdCompPago, doc.FEmision, 0,   Sum(mov.Debe) AS Debe, Sum(mov.Haber) AS Haber "
    Q1 = Q1 & " FROM ((Documento as doc   LEFT JOIN MovComprobante as mov ON doc.IdDoc = mov.IdDoc) "
    Q1 = Q1 & " LEFT JOIN Comprobante as com ON  com.idcomp = mov.idcomp ) "
    Q1 = Q1 & " where (com.Estado <> " & ED_ANULADO & "  Or com.Estado Is Null) "
    Q1 = Q1 & " AND  doc.TipoLib IN(" & LIB_OTROFULL & ") "
    Q1 = Q1 & " AND doc.SaldoDoc IS NULL "
    Q1 = Q1 & " AND doc.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND doc.Ano = " & Ano
    Q1 = Q1 & " GROUP BY doc.TipoLib, doc.IdDoc, doc.Total, doc.Estado, doc.TotpagadoAnoAnt, doc.IdDocAsoc, doc.Estado, doc.IdCompCent, doc.IdCompPago, doc.FEmision  ORDER BY  doc.IdDoc"


   
   Set Rs = OpenRs(DbMain, Q1)
      
      
   Do While Rs.EOF = False
                                
      
      'detalle doc
      If CurIdDoc <> vFld(Rs("IdDoc")) Then   ' puede venir más de una vez cuando hay documentos asociados
      
         If vFld(Rs("IdDoc")) = 483 Then   'OJO CON EL ESTADO DEL DOCUEMENTO DEL AñO ANTERIOR QUE DEBE ESTAR CENTRALIZADO O PAGADO
            MsgBeep vbExclamation
         End If

         Debe = 0
         Haber = 0
         SetPagado = False
         
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
         
            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
               Saldo = vFld(Rs("Total"))
               Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))
            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = Abs((Debe - Haber))
               Saldo = Abs((vFld(Rs("Total"))) - Abs(vFld(Rs("TotPagadoAnoAnt")))) - Saldo
               'Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))
            End If
            
            'Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            'Saldo = Abs((vFld(Rs("Total"))) - Abs(vFld(Rs("TotPagadoAnoAnt")))) - Saldo
         Else
         
         
'            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
'               Saldo = vFld(Rs("Total"))
'            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = Debe - Haber
            
               If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                  If Saldo <= 0 Then
                     DocOtroEsCargo = True
                  Else
                     DocOtroEsCargo = False
                  End If
               End If
            
'            End If
           
           
            If IsNull(Rs("TotPagadoAnoAnt")) = False Then    'FCA 04 feb 2020: se asume que sólo el primer año el TotPagadoAnoAnt es NULL, por lo tanto, al segundo año se le agrega el total
               Saldo = Saldo + IIf(vFld(Rs("DocOtroEsCargo")), -1, 1) * vFld(Rs("Total"))
            End If
            
           
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0 "
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              'Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
           
            Call ExecSQL(DbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
                              
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado, pero si ponemos el DocOtroEsCargo
                           
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Or Saldo = 0 Then
                     'Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               Else   'Saldo = Total
               
                  If vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano Then   'está asociado a un comprobante de centralización o es del año anterior
                     'Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     'Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
                                 
               End If
               
            Else
               Q1 = Q1 & ", DocOtroEsCargo = " & Abs(DocOtroEsCargo)
                                      
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
            Call ExecSQL(DbMain, Q1)
         
         End If
         
'      Else
'
'         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
'            Debe = Debe + vFld(Rs("Debe"))
'            Haber = Haber + vFld(Rs("Haber"))
'            Saldo = Debe - Haber
'
'            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))  'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'
'         Else
'            Debe = Debe + vFld(Rs("Debe"))
'            Haber = Haber + vFld(Rs("Haber"))
'            Saldo = Debe - Haber
'
'            If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
'                If Saldo <= 0 Then
'                   DocOtroEsCargo = True
'                Else
'                   DocOtroEsCargo = False
'                End If
'             End If
'
'            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'
'         End If
'
'
'
'         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
'            Q1 = "UPDATE DocumentoFull SET SaldoDoc = 0"
'
'            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
'              Q1 = Q1 & ", Estado = " & ED_PAGADO
'            End If
'
'            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
'            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
'
'            Call ExecSQL(DbMain, Q1)
'
'         Else
'            Q1 = "UPDATE DocumentoFull SET SaldoDoc = " & Saldo
'
'            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
'            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado
'
'               If Abs(Saldo) <> vFld(Rs("Total")) Then
'                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Or vFld(Rs("Estado")) = ED_PAGADO Then
'                     Q1 = Q1 & ", Estado = " & ED_PAGADO
'                     SetPagado = True
'                  End If
'
'               ElseIf Not SetPagado Then    'Saldo = Total
'
'                  If (vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano) Then    'está asociado a un comprobante de centralización o es del año anterior
'                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
'                  Else
'                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
'                  End If
'
'               End If
'            End If
'
'            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
'            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
'            Call ExecSQL(DbMain, Q1)
'         End If
         
      End If
         
      CurIdDoc = vFld(Rs("IdDoc"))
      ' 2828725 Cambia a estado pendiente los documentos NDV si los comprobantes no suman igual el debe y el haber
'        Q1 = "SELECT Switch(sum(debe) = sum(haber), 0,sum(debe) <> sum(haber),  abs(sum(debe) - sum(haber))) as pagado "
'        Q1 = Q1 & " FROM    documento as docu, movcomprobante as mov, comprobante com "
'        Q1 = Q1 & " WHERE   docu.iddoc = mov.iddoc "
'        Q1 = Q1 & " AND     mov.idcomp = com.idcomp "
'        Q1 = Q1 & " AND tipolib = 2 AND tipodoc = 4 "
'        Q1 = Q1 & " AND     docu.numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'        Set Rs2 = OpenRs(DbMain, Q1)
'
'
'        Do While Rs2.EOF = False
'
'            If vFld(Rs2("pagado")) > 0 Then
'
'            Q1 = "UPDATE documento "
'            Q1 = Q1 & " SET Estado = " & ED_PENDIENTE
'            Q1 = Q1 & " , SaldoDoc = " & vFld(Rs2("pagado"))
'            Q1 = Q1 & " WHERE numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'            Call ExecSQL(DbMain, Q1)
'
'            End If
'        Rs2.MoveNext
'        Loop
'        Call CloseRs(Rs2)
        ' FIN 2828725
               
      Rs.MoveNext
   Loop
      
   Call CloseRs(Rs)
   
End Function

Public Function GenDocsPendientes(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False, Optional ByVal ClearMsj As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String
   Dim ConnStr As String
   Dim Condicion As String
   Dim OtrosDocs As Long
   
   GenDocsPendientes = False
   
   If Not gEmprSeparadas Then
      GenDocsPendientes = GenDocsPendientesEmpJuntas(IdEmpresa, Rut, Ano, Msg, ClearFExported, ClearMsj)
      Exit Function
   End If
   
#If DATACON = 1 Then       'Access
   RutMdb = Rut & ".mdb"
   
   If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
      If Msg Then
       If ClearMsj = False Then
         MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar documentos pendientes en forma automática.", vbExclamation + vbOKOnly
       End If
      End If
      Exit Function
   End If
   
   If Not ClearFExported Then
     If ClearMsj = False Then
      If MsgBox1("¿Desea traer los documentos de los libros de Compras, Ventas y Retenciones desde el año anterior?" & vbCrLf & vbCrLf & "Solo se traerán aquellos que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
     End If
   End If
   
   'cerramos el año actual y abrimos el año anterior
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano - 1)
   Call LinkMdbAdm
   
   'corrige base del año anterior, por si las moscas
   Call CorrigeBase
   
   '14690904
   Call CloseDb(DbMain)
   Call OpenDbEmp(Rut, Ano - 1)
   '14690904
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   
  'If ClearMsj = False Then
   If FCierre = 0 Then
      If Msg Then
        If ClearMsj = False Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar documentos pendientes.", vbExclamation + vbOKOnly
        Else
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar cuadratura de documentos.", vbExclamation + vbOKOnly
         
        End If
      End If
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
      
      Exit Function
   End If
  'Else
   '   Call CloseDb(DbMain)
      'Call OpenDbEmp(Rut, Ano)
  'End If
   
   'recalculamos los saldos finales de los docs
   Call RecalcSaldos(gEmpresa.id, Ano - 1)
   Call RecalcSaldosFulle(gEmpresa.id, Ano - 1)
   
   
   'veamos si quedan docs del año anterior en Estado Pendiente, si es así, mandamos mensaje
   Q1 = "SELECT IdDoc FROM Documento WHERE Estado = " & ED_PENDIENTE & " AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If Msg Then
       If ClearMsj = False Then
         MsgBox1 "Hay documentos de Compras, Ventas o Retenciones del año anterior que se encuentran en estado Pendiente. " & vbNewLine & vbNewLine & "Sólo se traerán los documentos de los libros de Compras, Ventas y Retenciones, que están en estado Centralizado o Pagado y que tengan saldo distinto de cero.", vbExclamation + vbOKOnly
       End If
      End If
   End If
   
   Call CloseRs(Rs)
      
'   MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Sólo se traerán los documentos de los libros de Compras, Ventas " & vbCrLf & "y Retenciones, que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbExclamation + vbOKOnly
   
   If ClearFExported Then
      'Esta opción es sólo para el caso en que esté volviendo a generar el año
      'y haya borrado el archivo MDB del nuevo año por debajo.
      'Eso se detecta en ReadEmpresa
   
      Q1 = "UPDATE Documento SET FExported = 0 "
'      Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
      Call ExecSQL(DbMain, Q1)
   End If
   
   
   'copiamos los IdDoc e IdMovDoc para despues re-vincular las tablas
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND SaldoDoc <> 0 AND Estado IN(" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
  
'   'vemos si hay Otros Documentos en estado APROBADO para traer
'   'se eliminó esta funcionalidad de traer los otros docs porque hay muchos detalles que no están cubiertos para estos docs:
'   '    no aparecen en el analítico los del año anterior
'   '    se permite que se usen más de un}a vez
'   '    el estado lo maneja el usuario por lo que hay inconsistencias, etc.

'   Se volvió a agregar, manteniendo la funcionalidad actual pero agregando que el sistema calcule los saldos en base a los movimientos contables en los que se asocian los otros documentos (FCA 23 ene 2020)

   Q1 = "SELECT Count(*) FROM Documento "
   '616437 ffv odf
   'Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL  ) AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ") AND Estado IN(" & ED_APROBADO & "," & ED_CENTRALIZADO & "," & ED_PAGADO & ")"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL  ) AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ") AND Estado IN(" & ED_APROBADO & ")"
   '616437 ffv odf
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs(0)) > 0 Then
         If Msg Then
           If ClearMsj = False Then
              If MsgBox1("¿Desea traer Otros Documentos del año anterior que se encuentran en estado APROBADO?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
               'marcamos para exportar con -1
               Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
               Call ExecSQL(DbMain, Q1)
               
'                Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1" ' FExported > 0 2748525
'               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL Or FExported > 0 ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ")"
'               Call ExecSQL(DbMain, Q1)

               OtrosDocs = vFld(Rs(0))
              End If
              
             Else
             'marcamos para exportar con -1
               Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
               Call ExecSQL(DbMain, Q1)
               
'                Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1" ' FExported > 0 2748525
'               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL Or FExported > 0 ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ")"
'               Call ExecSQL(DbMain, Q1)

               OtrosDocs = vFld(Rs(0))

            End If
               
           
         End If
      End If
   End If

   Call CloseRs(Rs)
   
   'linkeamos la tabla de Documentos del año actual para agregar los Documentos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Documento", "DocumentoNew", , , gEmpresa.ConnStr)
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "MovDocumento", "MovDocumentoNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de Entidades del año actual para agregar nuevas entidades, si las hay
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Entidades", "EntidadesNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de Cuentas del año actual para hacer calzar los códigos si corresponde
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , gEmpresa.ConnStr)
   
   
   
   'Actualizamos los SALDOS de los documentos ya traidos en importaciones anteriores, antes de traer los nuevos documentos,
      
   'es Access
   'Compras, Ventas y Retenciones
   
   Condicion = "((Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ") AND TipoDocs.EsRebaja = 0) OR (Documento.TipoLib =" & LIB_VENTAS & " AND TipoDocs.EsRebaja <> 0)"
   
   Q1 = "UPDATE (DocumentoNew INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc )"
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " SET DocumentoNew.TotPagadoAnoAnt = "
   'Q1 = Q1 & " (iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 * Documento.Total, Documento.Total ), iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 *(Documento.Total - (-1 * Documento.SaldoDoc)), Documento.Total - Documento.SaldoDoc)))"
   Q1 = Q1 & " (iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, iif(" & Condicion & ", -1 * Documento.Total, Documento.Total ), iif(" & Condicion & ", -1 *(Documento.Total - (-1 * Documento.SaldoDoc)), Documento.Total - Documento.SaldoDoc)))"
   Q1 = Q1 & " WHERE Documento.FExported > 0 "
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'Otros Documentos
   
   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
   
   Q1 = "UPDATE (DocumentoNew INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc )"
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " SET DocumentoNew.TotPagadoAnoAnt = "
   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total - Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc))))"
   
   Q1 = Q1 & ", DocumentoNew.DocOtroEsCargo = Documento.DocOtroEsCargo "
   Q1 = Q1 & " WHERE Documento.FExported > 0 "
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
  
   
'    Q1 = "UPDATE (DocumentoNew INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc )"
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano )"
'   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
'   Q1 = Q1 & " SET DocumentoNew.TotPagadoAnoAnt = "
'   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total - Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc))))"
'
'   Q1 = Q1 & ", DocumentoNew.DocOtroEsCargo = Documento.DocOtroEsCargo "
'   Q1 = Q1 & " WHERE Documento.FExported > 0 "
'   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
''   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'   Call ExecSQL(DbMain, Q1)
   
   
   'primero traemos las entidades nuevas (dado los índices definidos en la tabla Entidades, sólo se insertarán los Ruts y Códigos nuevos (son únicos))
   '28 Sep 2006
   Q1 = "INSERT INTO EntidadesNew SELECT Entidades.* FROM Entidades"
   Q1 = Q1 & " WHERE Entidades.IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1)
      
   Where = " WHERE FExported < 0"
   
'   Q1 = "UPDATE Documento INNER JOIN DocumentoNew "
'   Q1 = Q1 & " ON Documento.TipoLib = DocumentoNew.TipoLib AND Documento.TipoDoc = DocumentoNew.TipoDoc AND Documento.NumDoc = DocumentoNew.NumDoc AND "
'
'   SET Documento.FExported = -2"
'
'
'
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan documentos repetidos
   Q1 = "UPDATE Documento INNER JOIN DocumentoNew ON Documento.idDoc = DocumentoNew.OldIdDoc "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano "
   Q1 = Q1 & " SET Documento.FExported=" & CLng(Int(Now))
   Q1 = Q1 & " WHERE Documento.FExported < 0"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'Agregamos docs marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   
   'Compras, Venyas y Retenciones
   Q1 = "INSERT INTO DocumentoNew "
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   'pipe bug
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld,DocOtroEsCargo,ValRet3Porc,IdCuentaRet3Porc "
   
   
   Q1 = Q1 & " FROM ((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
'   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = Documento.IdEmpresa AND Cuentas1.Ano = Documento.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
'   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = Documento.IdEmpresa AND Cuentas2.Ano = Documento.Ano "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision"
   Call ExecSQL(DbMain, Q1)
   
   'Otros Documentos
   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
   Q1 = "INSERT INTO DocumentoNew "
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
'   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total -  Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
   
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
      'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld " pipe
   'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico "
   
   'pipe bug
      Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico,ValRet3Porc,IdCuentaRet3Porc "
   
   Q1 = Q1 & " FROM ((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
'   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = Documento.IdEmpresa AND Cuentas1.Ano = Documento.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
'   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = Documento.IdEmpresa AND Cuentas2.Ano = Documento.Ano "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision"
   Call ExecSQL(DbMain, Q1)
   
   
   'otros documentos Full
   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
   Q1 = "INSERT INTO DocumentoNew "
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
'   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total -  Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
   
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
      'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld " pipe
   'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico "
   
   'pipe bug
      Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico,ValRet3Porc,IdCuentaRet3Porc, tratamiento "
   
   Q1 = Q1 & " FROM ((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
'   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = Documento.IdEmpresa AND Cuentas1.Ano = Documento.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
'   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = Documento.IdEmpresa AND Cuentas2.Ano = Documento.Ano "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_OTROFULL & ")"
   Q1 = Q1 & " AND (Documento.SaldoDoc IS NOT NULL AND Documento.SaldoDoc <> 0)"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision"
   Call ExecSQL(DbMain, Q1)

     
   'insertamos los MovDocumento, poniedo IdDoc en negativo (28 Sep 2006) para no confundirlo con algun MovDocuemtno que tiene el mismo IdDoc
   Q1 = "INSERT INTO MovDocumentoNew"
   Q1 = Q1 & " SELECT -1 * MovDocumento.IdDoc as IdDoc, 0 as IdCompCent, 0 as IdCompPago, MovDocumento.Orden, MovDocumento.IdCuenta, MovDocumento.Debe, MovDocumento.Haber, MovDocumento.Glosa, MovDocumento.IdTipoValLib, MovDocumento.EsTotalDoc, MovDocumento.IdCCosto, MovDocumento.IdAreaNeg, Cuentas.Codigo as CodCuentaOld, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ( MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "Documento") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   Q1 = Q1 & Where
''   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision "
   Call ExecSQL(DbMain, Q1)
   
   'Enlazamos MovDocumento a Documento, actualizando IdDoc, utilizando IdOldDoc
   Q1 = "UPDATE MovDocumentoNew INNER JOIN DocumentoNew ON -1 * MovDocumentoNew.IdDoc = DocumentoNew.OldIdDoc "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND DocumentoNew.Ano = MovDocumentoNew.Ano"
   Q1 = Q1 & " SET MovDocumentoNew.IdDoc = DocumentoNew.IdDoc "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   Q1 = "UPDATE (MovDocumentoNew INNER JOIN DocumentoNew ON MovDocumentoNew.IdDoc = DocumentoNew.IdDoc )"
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND DocumentoNew.Ano = MovDocumentoNew.Ano )"
   Q1 = Q1 & " INNER JOIN CuentasNew ON MovDocumentoNew.CodCuentaOld = CuentasNew.Codigo "
   Q1 = Q1 & " AND CuentasNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND CuentasNew.Ano = MovDocumentoNew.Ano"
   Q1 = Q1 & " SET MovDocumentoNew.IdCuenta = CuentasNew.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en MovDocumento
   Q1 = "UPDATE (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.CodCuentaOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   Q1 = Q1 & " SET MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & IdEmpresa & " AND MovDocumento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en Documento
   
   'Afecto
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaAfectoOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento")
   Q1 = Q1 & " SET Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'Exento
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaExentoOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas")
   Q1 = Q1 & " SET Documento.IdCuentaExento = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'Total
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaTotalOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas")
   Q1 = Q1 & " SET Documento.IdCuentaTotal = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'otros imp
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaOtroImp = IIF(TipoLib = " & LIB_VENTAS & ", " & gCtasBas.IdCtaOtrosImpDeb & ", " & gCtasBas.IdCtaOtrosImpCred & ")"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'IVA
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaIVA = IIF(TipoLib = " & LIB_VENTAS & ", " & gCtasBas.IdCtaIVADeb & ", " & gCtasBas.IdCtaIVACred & ")"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   
   'actualizamos el Id de la entidad en los documentos calzando por RUT
   Q1 = "UPDATE (DocumentoNew"
   Q1 = Q1 & " INNER JOIN Entidades ON DocumentoNew.IdEntidad = Entidades.IdEntidad ) "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
   Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
'   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
   Q1 = Q1 & " SET DocumentoNew.IdEntidad = EntidadesNew.IdEntidad "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   
   'pipe 2784017
    'actualizamos los campos ValRet3Porc - IdCuentaRet3Porc
   Q1 = "UPDATE (DocumentoNew"
   Q1 = Q1 & " INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc ) "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
 '  Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
'   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
   Q1 = Q1 & " SET DocumentoNew.ValRet3Porc = documento.ValRet3Porc "
    Q1 = Q1 & ", DocumentoNew.IdCuentaRet3Porc = documento.IdCuentaRet3Porc "
   Q1 = Q1 & " where "
   Q1 = Q1 & "  DocumentoNew.ValRet3Porc = 0 AND DocumentoNew.IdCuentaRet3Porc = 0 and   Documento.ValRet3Porc > 0 "
   Call ExecSQL(DbMain, Q1)

   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE DocumentoNew "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
      
   'Mensaje con cantidad Compras, Ventas o Retenciones
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
     If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " documentos de Compras, Ventas o Retenciones del año anterior, en estado Centralizado o Pagado, con saldo distinto de cero.", vbInformation
     End If
   End If
   
   Call CloseRs(Rs)
   
   'Mensaje con cantidad Otros Documentos
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
     If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " Otros Documentos del año anterior, en estado Aprobado, con saldo distinto de cero.", vbInformation
     End If
   End If
   
   Call CloseRs(Rs)
   
   'Mensaje con cantidad Otros Documentos
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_OTROFULL & ")"
   Q1 = Q1 & " AND (Documento.SaldoDoc IS NOT NULL AND Documento.SaldoDoc <> 0)"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
    If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " Otros Documentos Full del año anterior, en estado Aprobado, con saldo distinto de cero.", vbInformation
    End If
   End If
   
   Call CloseRs(Rs)
      
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'soltamos las tablas
   Q1 = "DROP TABLE DocumentoNew"
   Call ExecSQL(DbMain, Q1)

   Q1 = "DROP TABLE MovDocumentoNew"
   Call ExecSQL(DbMain, Q1)

   Q1 = "DROP TABLE EntidadesNew"
   Call ExecSQL(DbMain, Q1)
   
   'cerramos el año anterior y abrimos el año actual
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano)
   
   'recalculamos todos los saldos por si las moscas
   Q1 = "UPDATE Documento SET SaldoDoc = NULL "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
   Call ExecSQL(DbMain, Q1)
   
   Call RecalcSaldos(IdEmpresa, Ano)
   Call RecalcSaldosFulle(IdEmpresa, Ano)

   GenDocsPendientes = True

#End If

End Function

Public Function GenDocsFullPendientes(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String
   Dim ConnStr As String
   Dim Condicion As String
   Dim OtrosDocs As Long
   
   GenDocsFullPendientes = False
   
'   If Not gEmprSeparadas Then
'      GenDocsPendientes = GenDocsPendientesEmpJuntas(IdEmpresa, Rut, Ano, Msg, ClearFExported)
'      Exit Function
'   End If
   
#If DATACON = 1 Then       'Access
   RutMdb = Rut & ".mdb"
   
   
   If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
      If Msg Then
         MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar documentos pendientes en forma automática.", vbExclamation + vbOKOnly
      End If
      Exit Function
   End If
   
   If Not ClearFExported Then
      If MsgBox1("¿Desea traer los Otros documentos Full desde el año anterior?" & vbCrLf & vbCrLf & "Solo se traerán aquellos que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
   
   'cerramos el año actual y abrimos el año anterior
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano - 1)
   Call LinkMdbAdm
   
   'corrige base del año anterior, por si las moscas
   Call CorrigeBase
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
'      If Msg Then
'         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar documentos pendientes.", vbExclamation + vbOKOnly
'      End If
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
      
      Exit Function
   End If
   
   'recalculamos los saldos finales de los docs
   Call RecalcSaldos(gEmpresa.id, Ano - 1)
   Call RecalcSaldosFulle(gEmpresa.id, Ano - 1)
   
   
   'veamos si quedan docs del año anterior en Estado Pendiente, si es así, mandamos mensaje
   Q1 = "SELECT IdDoc FROM DocumentoFull WHERE Estado = " & ED_PENDIENTE & " AND TipoLib IN( " & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If Msg Then
         MsgBox1 "Hay documentos de Otros Documentos Full del año anterior que se encuentran en estado Pendiente. " & vbNewLine & vbNewLine & "Sólo se traerán los documentos del libro de Otros Documentos Full , que están en estado Pagado y que tengan saldo distinto de cero.", vbExclamation + vbOKOnly
      End If
   End If
   
   Call CloseRs(Rs)
      
'   MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Sólo se traerán los documentos de los libros de Compras, Ventas " & vbCrLf & "y Retenciones, que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbExclamation + vbOKOnly
   
   If ClearFExported Then
      'Esta opción es sólo para el caso en que esté volviendo a generar el año
      'y haya borrado el archivo MDB del nuevo año por debajo.
      'Eso se detecta en ReadEmpresa
   
      Q1 = "UPDATE DocumentoFull SET FExported = 0 "
'      Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
      Call ExecSQL(DbMain, Q1)
   End If
   
   
   'copiamos los IdDoc e IdMovDoc para despues re-vincular las tablas
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE DocumentoFull SET OldIdDocTmp = IdDoc, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND SaldoDoc <> 0 AND Estado IN(" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND TipoLib IN( " & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
  
'   'vemos si hay Otros Documentos en estado APROBADO para traer
'   'se eliminó esta funcionalidad de traer los otros docs porque hay muchos detalles que no están cubiertos para estos docs:
'   '    no aparecen en el analítico los del año anterior
'   '    se permite que se usen más de una vez
'   '    el estado lo maneja el usuario por lo que hay inconsistencias, etc.

'   Se volvió a agregar, manteniendo la funcionalidad actual pero agregando que el sistema calcule los saldos en base a los movimientos contables en los que se asocian los otros documentos (FCA 23 ene 2020)

'   Q1 = "SELECT Count(*) FROM Documento "
'   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL  ) AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ") AND Estado IN(" & ED_APROBADO & ")"
'
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'      If vFld(Rs(0)) > 0 Then
'         If Msg Then
'            If MsgBox1("¿Desea traer Otros Documentos del año anterior que se encuentran en estado APROBADO?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
'
'               'marcamos para exportar con -1
'               Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
'               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ")"
'               Call ExecSQL(DbMain, Q1)
'
''                Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1" ' FExported > 0 2748525
''               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL Or FExported > 0 ) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ")"
''               Call ExecSQL(DbMain, Q1)
'
'               OtrosDocs = vFld(Rs(0))
'
'            End If
'         End If
'      End If
'   End If
'
'   Call CloseRs(Rs)
   
   'linkeamos la tabla de Documentos del año actual para agregar los Documentos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "DocumentoFull", "DocumentoNew", , , gEmpresa.ConnStr)
   'Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "MovDocumentoFull", "MovDocumentoNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de Entidades del año actual para agregar nuevas entidades, si las hay
   'Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Entidades", "EntidadesNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de Cuentas del año actual para hacer calzar los códigos si corresponde
   'Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Cuentas", "CuentasNew", , , gEmpresa.ConnStr)
   
   
   
   'Actualizamos los SALDOS de los documentos ya traidos en importaciones anteriores, antes de traer los nuevos documentos,
      
   'es Access
   'Compras, Ventas y Retenciones
   
   Condicion = "((DocumentoFull.TipoLib =" & LIB_OTROFULL & ") AND TipoDocs.EsRebaja = 0) OR (DocumentoFull.TipoLib =" & LIB_OTROFULL & " AND TipoDocs.EsRebaja <> 0)"
   
   Q1 = "UPDATE (DocumentoNew INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc )"
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " SET DocumentoNew.TotPagadoAnoAnt = "
   'Q1 = Q1 & " (iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 * Documento.Total, Documento.Total ), iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 *(Documento.Total - (-1 * Documento.SaldoDoc)), Documento.Total - Documento.SaldoDoc)))"
   Q1 = Q1 & " (iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, iif(" & Condicion & ", -1 * Documento.Total, Documento.Total ), iif(" & Condicion & ", -1 *(Documento.Total - (-1 * Documento.SaldoDoc)), Documento.Total - Documento.SaldoDoc)))"
   Q1 = Q1 & " WHERE Documento.FExported > 0 "
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'Otros Documentos
   
'   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
'
'   Q1 = "UPDATE (DocumentoNew INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc )"
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano )"
'   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
'   Q1 = Q1 & " SET DocumentoNew.TotPagadoAnoAnt = "
'   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total - Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc))))"
'
'   Q1 = Q1 & ", DocumentoNew.DocOtroEsCargo = Documento.DocOtroEsCargo "
'   Q1 = Q1 & " WHERE Documento.FExported > 0 "
'   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
''   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
'   Call ExecSQL(DbMain, Q1)
   
   
   
   'primero traemos las entidades nuevas (dado los índices definidos en la tabla Entidades, sólo se insertarán los Ruts y Códigos nuevos (son únicos))
   '28 Sep 2006
'   Q1 = "INSERT INTO EntidadesNew SELECT Entidades.* FROM Entidades"
'   Q1 = Q1 & " WHERE Entidades.IdEmpresa = " & IdEmpresa
'   Call ExecSQL(DbMain, Q1)
      
   'Where = " WHERE FExported < 0"
   Where = " "
   
'   Q1 = "UPDATE Documento INNER JOIN DocumentoNew "
'   Q1 = Q1 & " ON Documento.TipoLib = DocumentoNew.TipoLib AND Documento.TipoDoc = DocumentoNew.TipoDoc AND Documento.NumDoc = DocumentoNew.NumDoc AND "
'
'   SET Documento.FExported = -2"
'
'
'
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan documentos repetidos
   Q1 = "UPDATE Documento INNER JOIN DocumentoNew ON Documento.idDoc = DocumentoNew.OldIdDoc "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano "
   Q1 = Q1 & " SET Documento.FExported=" & CLng(Int(Now))
   Q1 = Q1 & " WHERE Documento.FExported < 0"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'Agregamos docs marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   
   'Compras, Venyas y Retenciones
   Q1 = "INSERT INTO DocumentoNew "
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc < 0,sum(-1*(movc.Debe + movc.Haber)),sum(movc.Debe + movc.Haber)) as  TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   'pipe bug
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld,DocOtroEsCargo,ValRet3Porc,IdCuentaRet3Porc "
   
   
   Q1 = Q1 & " FROM (((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
'   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = Documento.IdEmpresa AND Cuentas1.Ano = Documento.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta) "
   Q1 = Q1 & " LEFT JOIN MovComprobanteFull movc ON movc.IdDoc = DocumentoFull.IdDoc "
'   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = Documento.IdEmpresa AND Cuentas2.Ano = Documento.Ano "
   'Q1 = Q1 & Where
   Q1 = Q1 & " WHERE Documento.TipoLib IN( " & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision"
   Q1 = Q1 & " GROUP BY IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal, "
   Q1 = Q1 & IdEmpresa & ", " & Ano & ",  Cuentas.Codigo, Cuentas1.Codigo, Cuentas2.Codigo,DocOtroEsCargo,ValRet3Porc,IdCuentaRet3Porc "
   
   
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
'   'Otros Documentos
'   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
'   Q1 = "INSERT INTO DocumentoNew "
'   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
''   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
'   Q1 = Q1 & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total -  Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
'
'   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
'      'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld " pipe
'   'Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico "
'
'   'pipe bug
'      Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico,ValRet3Porc,IdCuentaRet3Porc "
'
'   Q1 = Q1 & " FROM ((Documento "
'   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento") & " )"
'   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
''   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = Documento.IdEmpresa AND Cuentas1.Ano = Documento.Ano )"
'   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
''   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = Documento.IdEmpresa AND Cuentas2.Ano = Documento.Ano "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
''   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'   'Q1 = Q1 & " ORDER BY Documento.FEmision"
'   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
'   Call ExecSQL(DbMain, Q1)
   

     
   'insertamos los MovDocumento, poniedo IdDoc en negativo (28 Sep 2006) para no confundirlo con algun MovDocuemtno que tiene el mismo IdDoc
'   Q1 = "INSERT INTO MovDocumentoNew"
'   Q1 = Q1 & " SELECT -1 * MovDocumento.IdDoc as IdDoc, 0 as IdCompCent, 0 as IdCompPago, MovDocumento.Orden, MovDocumento.IdCuenta, MovDocumento.Debe, MovDocumento.Haber, MovDocumento.Glosa, MovDocumento.IdTipoValLib, MovDocumento.EsTotalDoc, MovDocumento.IdCCosto, MovDocumento.IdAreaNeg, Cuentas.Codigo as CodCuentaOld, "
'   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano "
'   Q1 = Q1 & " FROM ( MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "Documento") & " )"
'   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
'   Q1 = Q1 & Where
'''   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'   'Q1 = Q1 & " ORDER BY Documento.FEmision "
'   Q1 = Replace(Replace(Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull."), "MovDocumento.", "MovDocumentoFull."), "MovDocumento ", "MovDocumentoFull ")
'   Call ExecSQL(DbMain, Q1)
   
   'Enlazamos MovDocumento a Documento, actualizando IdDoc, utilizando IdOldDoc
'   Q1 = "UPDATE MovDocumentoNew INNER JOIN DocumentoNew ON -1 * MovDocumentoNew.IdDoc = DocumentoNew.OldIdDoc "
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND DocumentoNew.Ano = MovDocumentoNew.Ano"
'   Q1 = Q1 & " SET MovDocumentoNew.IdDoc = DocumentoNew.IdDoc "
'   Q1 = Q1 & Where
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
'   Q1 = Replace(Replace(Q1, "MovDocumento ", "MovDocumentoFull "), "MovDocumento.", "MovDocumentoFull.")
'   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
'   Q1 = "UPDATE (MovDocumentoNew INNER JOIN DocumentoNew ON MovDocumentoNew.IdDoc = DocumentoNew.IdDoc )"
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND DocumentoNew.Ano = MovDocumentoNew.Ano )"
'   Q1 = Q1 & " INNER JOIN CuentasNew ON MovDocumentoNew.CodCuentaOld = CuentasNew.Codigo "
'   Q1 = Q1 & " AND CuentasNew.IdEmpresa = MovDocumentoNew.IdEmpresa AND CuentasNew.Ano = MovDocumentoNew.Ano"
'   Q1 = Q1 & " SET MovDocumentoNew.IdCuenta = CuentasNew.IdCuenta "
'   Q1 = Q1 & Where
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
'   Q1 = Replace(Replace(Q1, "MovDocumento ", "MovDocumentoFull "), "MovDocumento.", "MovDocumentoFull.")
'   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en MovDocumento
'   Q1 = "UPDATE (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
'   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.CodCuentaOld = Cuentas.Codigo "
'   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
'   Q1 = Q1 & " SET MovDocumento.IdCuenta = Cuentas.IdCuenta "
'   Q1 = Q1 & Where
''   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & IdEmpresa & " AND MovDocumento.Ano = " & Ano
'   Q1 = Replace(Replace(Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull."), "MovDocumento.", "MovDocumentoFull."), "MovDocumento ", "MovDocumentoFull ")
'   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en Documento
   
   'Afecto
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaAfectoOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "Documento")
   Q1 = Q1 & " SET Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'Exento
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaExentoOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas")
   Q1 = Q1 & " SET Documento.IdCuentaExento = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'Total
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaTotalOld = Cuentas.Codigo "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Cuentas")
   Q1 = Q1 & " SET Documento.IdCuentaTotal = Cuentas.IdCuenta "
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'otros imp
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaOtroImp = IIF(TipoLib = " & LIB_OTROFULL & ", " & gCtasBas.IdCtaOtrosImpDeb & ", " & gCtasBas.IdCtaOtrosImpCred & ")"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'IVA
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaIVA = IIF(TipoLib = " & LIB_OTROFULL & ", " & gCtasBas.IdCtaIVADeb & ", " & gCtasBas.IdCtaIVACred & ")"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   
   'actualizamos el Id de la entidad en los documentos calzando por RUT
'   Q1 = "UPDATE (DocumentoNew"
'   Q1 = Q1 & " INNER JOIN Entidades ON DocumentoNew.IdEntidad = Entidades.IdEntidad ) "
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
'   Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
''   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
'   Q1 = Q1 & " SET DocumentoNew.IdEntidad = EntidadesNew.IdEntidad "
'   Q1 = Q1 & Where
''   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   
   'pipe 2784017
    'actualizamos los campos ValRet3Porc - IdCuentaRet3Porc
   Q1 = "UPDATE (DocumentoNew"
   Q1 = Q1 & " INNER JOIN Documento ON DocumentoNew.OldIdDoc = Documento.IdDoc ) "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
 '  Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
'   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
   Q1 = Q1 & " SET DocumentoNew.ValRet3Porc = Documento.ValRet3Porc "
    Q1 = Q1 & ", DocumentoNew.IdCuentaRet3Porc = Documento.IdCuentaRet3Porc "
   Q1 = Q1 & " where "
   Q1 = Q1 & "  DocumentoNew.ValRet3Porc = 0 AND DocumentoNew.IdCuentaRet3Porc = 0 and   Documento.ValRet3Porc > 0 "
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   
   
   
   
   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE DocumentoNew "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
      
   'Mensaje con cantidad Compras, Ventas o Retenciones
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " WHERE Documento.TipoLib IN( " & LIB_OTROFULL & ")"
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " documentos de Otros Documentos Full del año anterior, en estado Pagado, con saldo distinto de cero.", vbInformation
   End If
   
   Call CloseRs(Rs)
   
'   'Mensaje con cantidad Otros Documentos
'   Q1 = "SELECT Count(*) As N "
'   Q1 = Q1 & " FROM Documento "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
''   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Rs.EOF = False Then
'      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " Otros Documentos del año anterior, en estado Aprobado, con saldo distinto de cero.", vbInformation
'   End If
   
'   Call CloseRs(Rs)
      
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
'   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   'soltamos las tablas
   Q1 = "DROP TABLE DocumentoNew"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "DROP TABLE MovDocumentoNew"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "DROP TABLE EntidadesNew"
   Call ExecSQL(DbMain, Q1)
   
   'cerramos el año anterior y abrimos el año actual
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano)
   
   'recalculamos todos los saldos por si las moscas
   Q1 = "UPDATE Documento SET SaldoDoc = NULL "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_OTROFULL & ")"
   Q1 = Replace(Replace(Q1, "Documento ", "DocumentoFull "), "Documento.", "DocumentoFull.")
   Call ExecSQL(DbMain, Q1)
   
   Call RecalcSaldos(IdEmpresa, Ano)
   Call RecalcSaldosFulle(IdEmpresa, Ano)

   GenDocsFullPendientes = True

#End If

End Function

'genera los docs que quedaron pendientes del año anterior para Base con Empresas Juntas
Public Function GenDocsPendientesEmpJuntas(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False, Optional ByVal ClearMsj As Boolean = False) As Boolean

#If DATACON = 2 Then       'SQL Server o MySQL
   
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String, WhereOld As String
   Dim ConnStr As String
   Dim Condicion As String
   Dim OtrosDocs As Long
   Dim OtrosDocsFull As Long
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
      
   
   GenDocsPendientesEmpJuntas = False
   
   If gEmprSeparadas Then
      Exit Function
   End If

   If gEmpresa.TieneAnoAntAccess Then  'los docs pendientes ya fueron generados al crear el nuevo año desde Access
      GenDocsPendientesEmpJuntas = True
      Exit Function
   End If
   
   If Not ClearFExported Then
    If ClearMsj = False Then
      If MsgBox1("¿Desea traer los documentos de los libros de Compras, Ventas y Retenciones desde el año anterior?" & vbCrLf & vbCrLf & "Solo se traerán aquellos que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
    End If
   End If
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   Else
    If ClearMsj = False Then
      MsgBox1 "No hay registro de año anterior. No es posible generar documentos pendientes.", vbExclamation + vbOKOnly
    End If
      Call CloseRs(Rs)
      Exit Function
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
       If ClearMsj = False Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible generar documentos pendientes.", vbExclamation + vbOKOnly
       End If
      End If
      
      Exit Function
   End If
   
   'recalculamos los saldos finales de los docs
   Call RecalcSaldos(gEmpresa.id, Ano - 1)
   Call RecalcSaldosFulle(gEmpresa.id, Ano - 1)
   
   
   'veamos si quedan docs del año anterior en Estado Pendiente, si es así, mandamos mensaje
   Q1 = "SELECT IdDoc FROM Documento WHERE Estado = " & ED_PENDIENTE & " AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If Msg Then
       If ClearMsj = False Then
         MsgBox1 "Hay documentos de Compras, Ventas o Retenciones del año anterior que se encuentran en estado Pendiente. " & vbNewLine & vbNewLine & "Sólo se traerán los documentos de los libros de Compras, Ventas y Retenciones, que están en estado Centralizado o Pagado y que tengan saldo distinto de cero.", vbExclamation + vbOKOnly
       End If
      End If
   End If
   
   Call CloseRs(Rs)
      
'   MsgBox1 "Atención:" & vbNewLine & vbNewLine & "Sólo se traerán los documentos de los libros de Compras, Ventas " & vbCrLf & "y Retenciones, que están en estado Centralizado o Pagado, que tengan saldo distinto de cero, y que no hayan sido importados con anterioridad.", vbExclamation + vbOKOnly
   
   If ClearFExported Then
      'Esta opción es sólo para el caso en que esté volviendo a generar el año
      'y haya borrado el archivo MDB del nuevo año por debajo.
      'Eso se detecta en ReadEmpresa
   
      Q1 = "UPDATE Documento SET FExported = 0 "
      Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
      Call ExecSQL(DbMain, Q1)
   End If
   
   
   'copiamos los IdDoc e IdMovDoc para despues re-vincular las tablas
   
   
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND SaldoDoc <> 0 AND Estado IN(" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_OTROFULL & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
  
'   'vemos si hay Otros Documentos en estado APROBADO para traer
'   'se eliminó esta funcionalidad de traer los otros docs porque hay muchos detalles que no están cubiertos para estos docs:
'   '    no aparecen en el analítico los del año anterior
'   '    se permite que se usen más de una vez
'   '    el estado lo maneja el usuario por lo que hay inconsistencias, etc.

'   Se volvió a agregar, manteniendo la funcionalidad actual pero agregando que el sistema calcule los saldos en base a los movimientos contables en los que se asocian los otros documentos (FCA 23 ene 2020)

   Q1 = "SELECT Count(*) FROM Documento "
   
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND TipoLib IN (" & LIB_REMU & "," & LIB_OTROS & ") AND Estado IN(" & ED_APROBADO & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1

   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs(0)) > 0 Then
         If Msg Then
          If ClearMsj = False Then
            If MsgBox1("¿Desea traer Otros Documentos del año anterior que se encuentran en estado APROBADO?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
          End If

               'marcamos para exportar con -1
               Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_REMU & "," & LIB_OTROS & ")"
               Call ExecSQL(DbMain, Q1)

               OtrosDocs = vFld(Rs(0))

            End If
         End If
      End If
   End If

   Call CloseRs(Rs)

    'Otros documentos full
     'traspaso sql ODF ffv
   Q1 = "SELECT Count(*) FROM Documento "
   'Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND TipoLib IN (" & LIB_REMU & "," & LIB_OTROS & ") AND Estado IN(" & ED_APROBADO & ")"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND TipoLib IN (" & LIB_OTROFULL & ") AND Estado IN(" & ED_APROBADO & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1

   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs(0)) > 0 Then
         If Msg Then
           If ClearMsj = False Then
            If MsgBox1("¿Desea traer Otros Documentos full del año anterior que se encuentran en estado APROBADO?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
           End If

               'marcamos para exportar con -1
               Q1 = "UPDATE Documento SET OldIdDocTmp = IdDoc, FExported = -1"
               Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND SaldoDoc <> 0 AND Estado IN(" & ED_APROBADO & ") AND TipoLib IN ( " & LIB_OTROFULL & ")"
               Call ExecSQL(DbMain, Q1)

               OtrosDocsFull = vFld(Rs(0))

            End If
         End If
      End If
   End If

   Call CloseRs(Rs)
      
   'Antes de traer los nuevos documentos, actualizamos los saldos de los documentos ya traidos en importaciones anteriores
   
   'Ventas, Compras y Retenciones
   Condicion = "((DocumentoOld.TipoLib =" & LIB_COMPRAS & " or DocumentoOld.TipoLib = " & LIB_RETEN & ") AND TipoDocs.EsRebaja = 0) OR (DocumentoOld.TipoLib =" & LIB_VENTAS & " AND TipoDocs.EsRebaja <> 0)"
   
   Tbl = " Documento "
   sFrom = " (Documento INNER JOIN Documento As DocumentoOld ON Documento.OldIdDoc = DocumentoOld.IdDoc "
   sFrom = sFrom & " AND Documento.IdEmpresa = DocumentoOld.IdEmpresa AND Documento.Ano - 1 = DocumentoOld.Ano )"
   sFrom = sFrom & " INNER JOIN TipoDocs ON DocumentoOld.TipoLib = TipoDocs.TipoLib AND DocumentoOld.TipoDoc = TipoDocs.TipoDoc "
   sSet = " Documento.TotPagadoAnoAnt = "
   'Q1 = Q1 & " (iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 * Documento.Total, Documento.Total ), iif(Documento.TipoLib =" & LIB_COMPRAS & " or Documento.TipoLib = " & LIB_RETEN & ", -1 *(Documento.Total - (-1 * Documento.SaldoDoc)), Documento.Total - Documento.SaldoDoc)))"
   sSet = sSet & " (iif(DocumentoOld.SaldoDoc IS NULL or DocumentoOld.SaldoDoc = 0, iif(" & Condicion & ", -1 * DocumentoOld.Total, DocumentoOld.Total ), iif(" & Condicion & ", -1 *(DocumentoOld.Total - (-1 * DocumentoOld.SaldoDoc)), DocumentoOld.Total - DocumentoOld.SaldoDoc)))"
   sWhere = " WHERE DocumentoOld.FExported > 0 AND DocumentoOld.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   sWhere = sWhere & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Otros Documentos
   Condicion = " Documento.DocOtroEsCargo IS NULL OR Documento.DocOtroEsCargo = 0 "
   
   Tbl = " Documento "
   sFrom = " (Documento INNER JOIN Documento As DocumentoOld ON Documento.OldIdDoc = DocumentoOld.IdDoc "
   sFrom = sFrom & " AND Documento.IdEmpresa = DocumentoOld.IdEmpresa AND Documento.Ano - 1 = DocumentoOld.Ano )"
   sFrom = sFrom & " INNER JOIN TipoDocs ON DocumentoOld.TipoLib = TipoDocs.TipoLib AND DocumentoOld.TipoDoc = TipoDocs.TipoDoc "
   sSet = " Documento.TotPagadoAnoAnt = "
   sSet = sSet & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total - Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc))))"
   sSet = sSet & ", Documento.DocOtroEsCargo = Documento.DocOtroEsCargo "
   sWhere = " WHERE DocumentoOld.FExported > 0 AND DocumentoOld.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
   sWhere = sWhere & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Otros documentos full
   'traspaso odf ffv
   Tbl = " Documento "
   sFrom = " (Documento INNER JOIN Documento As DocumentoOld ON Documento.OldIdDoc = DocumentoOld.IdDoc "
   sFrom = sFrom & " AND Documento.IdEmpresa = DocumentoOld.IdEmpresa AND Documento.Ano - 1 = DocumentoOld.Ano )"
   sFrom = sFrom & " INNER JOIN TipoDocs ON DocumentoOld.TipoLib = TipoDocs.TipoLib AND DocumentoOld.TipoDoc = TipoDocs.TipoDoc "
   sSet = " Documento.TotPagadoAnoAnt = "
   sSet = sSet & " iif(Documento.SaldoDoc IS NULL or Documento.SaldoDoc = 0, 0, iif(" & Condicion & ", Documento.Total - Documento.SaldoDoc, -1 * (Documento.Total - (-1 * Documento.SaldoDoc))))"
   sSet = sSet & ", Documento.DocOtroEsCargo = Documento.DocOtroEsCargo "
   sWhere = " WHERE DocumentoOld.FExported > 0 AND DocumentoOld.TipoLib IN( " & LIB_OTROFULL & ")"
   sWhere = sWhere & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

   
   Where = " WHERE Documento.FExported < 0"
   WhereOld = " WHERE DocumentoOld.FExported < 0"
   
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan documentos repetidos
'   Q1 = "UPDATE Documento INNER JOIN Documento as DocumentoNew ON Documento.idDoc = DocumentoNew.OldIdDoc "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano "
'   Q1 = Q1 & " SET Documento.FExported=" & CLng(Int(Now))
'   Q1 = Q1 & " WHERE Documento.FExported < 0"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento  "
   sFrom = " Documento INNER JOIN Documento as DocumentoNew ON Documento.idDoc = DocumentoNew.OldIdDoc "
   sFrom = sFrom & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano "
   sSet = " Documento.FExported=" & CLng(Int(Now))
   sWhere = " WHERE Documento.FExported < 0"
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
'3026009
   sWhere = sWhere & " AND Documento.TipoLib <> " & LIB_OTROFULL
   '3026009
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Agregamos docs marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   
   'Compras, Ventas y Retenciones
   Q1 = "INSERT INTO Documento "
   Q1 = Q1 & " (IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " TotPagadoAnoAnt, "
   Q1 = Q1 & " IdEmpresa, Ano, "
   Q1 = Q1 & " CodCtaAfectoOld, CodCtaExentoOld, CodCtaTotalOld )"
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DocumentoOld.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   Q1 = Q1 & " FROM ((Documento as DocumentoOld"
   Q1 = Q1 & " LEFT JOIN Cuentas ON DocumentoOld.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & "  AND Cuentas.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON DocumentoOld.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas1.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON DocumentoOld.IdCuentaTotal = Cuentas2.IdCuenta "
   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas2.Ano = DocumentoOld.Ano "
   Q1 = Q1 & WhereOld
   Q1 = Q1 & " AND DocumentoOld.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY DocumentoOld.FEmision"
   Call ExecSQL(DbMain, Q1)
   
   
   'Otros Documentos
   
   Condicion = " DocumentoOld.DocOtroEsCargo IS NULL OR DocumentoOld.DocOtroEsCargo = 0 "
   
   Q1 = "INSERT INTO Documento "
   Q1 = Q1 & " (IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " TotPagadoAnoAnt, DocOtroEsCargo, "
   Q1 = Q1 & " IdEmpresa, Ano, "
   Q1 = Q1 & " CodCtaAfectoOld, CodCtaExentoOld, CodCtaTotalOld )"
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DocumentoOld.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(" & Condicion & ", Total - (-1 * SaldoDoc), -1 * (Total - (-1 * SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   Q1 = Q1 & " FROM ((Documento as DocumentoOld"
   Q1 = Q1 & " LEFT JOIN Cuentas ON DocumentoOld.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & "  AND Cuentas.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON DocumentoOld.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas1.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON DocumentoOld.IdCuentaTotal = Cuentas2.IdCuenta "
   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas2.Ano = DocumentoOld.Ano "
   Q1 = Q1 & WhereOld
   Q1 = Q1 & " AND DocumentoOld.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
   Q1 = Q1 & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY DocumentoOld.FEmision"
   Call ExecSQL(DbMain, Q1)
   
    'Otros Documentos full
   
   Condicion = " DocumentoOld.DocOtroEsCargo IS NULL OR DocumentoOld.DocOtroEsCargo = 0 "
   Q1 = ""
   Q1 = "INSERT INTO Documento "
   Q1 = Q1 & " (IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " TotPagadoAnoAnt, DocOtroEsCargo, "
   Q1 = Q1 & " IdEmpresa, Ano, "
   Q1 = Q1 & " CodCtaAfectoOld, CodCtaExentoOld, CodCtaTotalOld ,DocOtrosEnAnalitico,ValRet3Porc,IdCuentaRet3Porc, tratamiento )"
'   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DocumentoOld.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
'   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(" & Condicion & ", Total - (-1 * SaldoDoc), -1 * (Total - (-1 * SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
'   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
'   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
    Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DocumentoOld.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(DocumentoOld.SaldoDoc IS NULL or DocumentoOld.SaldoDoc = 0, 0, iif(" & Condicion & ", DocumentoOld.Total -  DocumentoOld.SaldoDoc, -1 * (DocumentoOld.Total - (-1 * DocumentoOld.SaldoDoc)))) As TotPagadoAnoAnt, DocOtroEsCargo, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   'pipe bug
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld   ,DocOtrosEnAnalitico,ValRet3Porc,IdCuentaRet3Porc, tratamiento "
   
   Q1 = Q1 & " FROM ((Documento as DocumentoOld"
   Q1 = Q1 & " LEFT JOIN Cuentas ON DocumentoOld.IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & "  AND Cuentas.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON DocumentoOld.IdCuentaExento = Cuentas1.IdCuenta "
   Q1 = Q1 & "  AND Cuentas1.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas1.Ano = DocumentoOld.Ano )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON DocumentoOld.IdCuentaTotal = Cuentas2.IdCuenta "
   Q1 = Q1 & "  AND Cuentas2.IdEmpresa = DocumentoOld.IdEmpresa AND Cuentas2.Ano = DocumentoOld.Ano "
   Q1 = Q1 & WhereOld
   Q1 = Q1 & " AND DocumentoOld.TipoLib IN( " & LIB_OTROFULL & ")"
   Q1 = Q1 & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY DocumentoOld.FEmision"
   Call ExecSQL(DbMain, Q1)

   
   'insertamos los MovDocumento, poniendo IdDoc en negativo (28 Sep 2006) para no confundirlo con algun MovDocuemtno que tiene el mismo IdDoc
   Q1 = "INSERT INTO MovDocumento "
   Q1 = Q1 & " ( IdDoc, IdCompCent, IdCompPago, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, CodCuentaOld, "
   Q1 = Q1 & " IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT -1 * MovDocumentoOld.IdDoc as IdDoc, 0 as IdCompCent, 0 as IdCompPago, MovDocumentoOld.Orden, MovDocumentoOld.IdCuenta, MovDocumentoOld.Debe, MovDocumentoOld.Haber, MovDocumentoOld.Glosa, MovDocumentoOld.IdTipoValLib, MovDocumentoOld.EsTotalDoc, MovDocumentoOld.IdCCosto, MovDocumentoOld.IdAreaNeg, CuentasOld.Codigo as CodCuentaOld, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ( MovDocumento As MovDocumentoOld INNER JOIN Documento as DocumentoOld ON MovDocumentoOld.IdDoc = DocumentoOld.IdDoc "
   Q1 = Q1 & "  AND DocumentoOld.IdEmpresa = MovDocumentoOld.IdEmpresa AND DocumentoOld.Ano = MovDocumentoOld.Ano ) "
   Q1 = Q1 & " INNER JOIN Cuentas as CuentasOld ON MovDocumentoOld.IdCuenta = CuentasOld.IdCuenta "
   Q1 = Q1 & "  AND CuentasOld.IdEmpresa = MovDocumentoOld.IdEmpresa AND CuentasOld.Ano = MovDocumentoOld.Ano"
   Q1 = Q1 & WhereOld
   Q1 = Q1 & " AND DocumentoOld.IdEmpresa = " & IdEmpresa & " AND DocumentoOld.Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision "
   Call ExecSQL(DbMain, Q1)
   
   'Enlazamos MovDocumento a Documento, actualizando IdDoc, utilizando IdOldDoc
'   Q1 = "UPDATE MovDocumento INNER JOIN Documento ON -1 * MovDocumento.IdDoc = Documento.OldIdDoc "
'   Q1 = Q1 & " AND Documento.IdEmpresa = MovDocumento.IdEmpresa AND Documento.Ano = MovDocumento.Ano"
'   Q1 = Q1 & " SET MovDocumento.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " MovDocumento "
   sFrom = " MovDocumento INNER JOIN Documento ON -1 * MovDocumento.IdDoc = Documento.OldIdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "MovDocumento")
   sSet = " MovDocumento.IdDoc = Documento.IdDoc "
   sWhere = Where
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en MovDocumento
'   Q1 = "UPDATE (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
'   Q1 = Q1 & "  AND Documento.IdEmpresa = MovDocumento.IdEmpresa AND Documento.Ano = MovDocumento.Ano )"
'   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.CodCuentaOld = Cuentas.Codigo "
'   Q1 = Q1 & "  AND Cuentas.IdEmpresa = MovDocumento.IdEmpresa AND Cuentas.Ano = MovDocumento.Ano"
'   Q1 = Q1 & " SET MovDocumento.IdCuenta = Cuentas.IdCuenta "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & IdEmpresa & " AND MovDocumento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " MovDocumento "
   sFrom = " (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.CodCuentaOld = Cuentas.Codigo "
   sFrom = sFrom & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   sSet = " MovDocumento.IdCuenta = Cuentas.IdCuenta "
   sWhere = Where
   sWhere = sWhere & " AND MovDocumento.IdEmpresa = " & IdEmpresa & " AND MovDocumento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Actualizamos IdCuenta por si el IdCuenta/Código del año actual no coicide con el del año anterior
   'en Documento
   'Afecto
'   Q1 = "UPDATE Documento "
'   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaAfectoOld = Cuentas.Codigo "
'   Q1 = Q1 & "  AND Documento.IdEmpresa = Cuentas.IdEmpresa AND Documento.Ano = Cuentas.Ano"
'   Q1 = Q1 & " SET Documento.IdCuentaAfecto = Cuentas.IdCuenta "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   sFrom = " Documento "
   sFrom = sFrom & " INNER JOIN Cuentas ON Documento.CodCtaAfectoOld = Cuentas.Codigo "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Cuentas")
   sSet = " Documento.IdCuentaAfecto = Cuentas.IdCuenta "
   sWhere = Where
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Exento
'   Q1 = "UPDATE Documento "
'   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaExentoOld = Cuentas.Codigo "
'   Q1 = Q1 & "  AND Documento.IdEmpresa = Cuentas.IdEmpresa AND Documento.Ano = Cuentas.Ano"
'   Q1 = Q1 & " SET Documento.IdCuentaExento = Cuentas.IdCuenta "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   sFrom = " Documento "
   sFrom = sFrom & " INNER JOIN Cuentas ON Documento.CodCtaExentoOld = Cuentas.Codigo "
   sFrom = sFrom & JoinEmpAno(gDbType, "Cuentas", "Documento")
   sSet = " Documento.IdCuentaExento = Cuentas.IdCuenta "
   sWhere = Where
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Total
'   Q1 = "UPDATE Documento "
'   Q1 = Q1 & " INNER JOIN Cuentas ON Documento.CodCtaTotalOld = Cuentas.Codigo "
'   Q1 = Q1 & "  AND Documento.IdEmpresa = Cuentas.IdEmpresa AND Documento.Ano = Cuentas.Ano"
'   Q1 = Q1 & " SET Documento.IdCuentaTotal = Cuentas.IdCuenta "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   sFrom = " Documento "
   sFrom = sFrom & " INNER JOIN Cuentas ON Documento.CodCtaTotalOld = Cuentas.Codigo "
   sFrom = sFrom & JoinEmpAno(gDbType, "Cuentas", "Documento")
   sSet = " Documento.IdCuentaTotal = Cuentas.IdCuenta "
   sWhere = Where
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'otros imp
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaOtroImp = IIF(TipoLib = " & LIB_VENTAS & ", " & gCtasBas.IdCtaOtrosImpDeb & ", " & gCtasBas.IdCtaOtrosImpCred & ")"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'IVA
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET Documento.IdCuentaIVA = IIF(TipoLib = " & LIB_VENTAS & ", " & gCtasBas.IdCtaIVADeb & ", " & gCtasBas.IdCtaIVACred & ")"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos el Id de la entidad en los documentos calzando por RUT
'   Q1 = "UPDATE (DocumentoNew"
'   Q1 = Q1 & " INNER JOIN Entidades ON DocumentoNew.IdEntidad = Entidades.IdEntidad "
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
'   Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
'   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
'   Q1 = Q1 & " SET DocumentoNew.IdEntidad = EntidadesNew.IdEntidad "
'   Q1 = Q1 & Where
'   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)


   'otros documentos full
   'traspaso odf ffv
   Tbl = " Documento "
   sFrom = " Documento "
   sFrom = sFrom & " INNER JOIN Cuentas ON Documento.IdCtaBanco = Cuentas.IdCuentaOld "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Cuentas")
   sSet = " Documento.IdCtaBanco = Cuentas.IdCuenta "
   sWhere = Where
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
      
   'Mensaje con cantidad Compras, Ventas o Retenciones
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
    If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " documentos de Compras, Ventas o Retenciones del año anterior, en estado Centralizado o Pagado, con saldo distinto de cero.", vbInformation
    End If
   End If
   
   Call CloseRs(Rs)
   
   'Mensaje con cantidad Otros Documentos
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_REMU & "," & LIB_OTROS & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
     If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " Otros Documentos del año anterior, en estado Aprobado, con saldo distinto de cero.", vbInformation
     End If
   End If
   
   Call CloseRs(Rs)
   
   'Mensaje con cantidad Otros Documentos full
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_OTROFULL & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
    If ClearMsj = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " Otros Documentos full del año anterior, en estado Aprobado, con saldo distinto de cero.", vbInformation
    End If
   End If
   
   Call CloseRs(Rs)
      
   
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
      
   'recalculamos todos los saldos por si las moscas
   Q1 = "UPDATE Documento SET SaldoDoc = NULL "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & " AND Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & "," & LIB_OTROFULL & ")"
   Call ExecSQL(DbMain, Q1)
   
   
   Call RecalcSaldos(IdEmpresa, Ano)
   Call RecalcSaldosFulle(IdEmpresa, Ano)

   GenDocsPendientesEmpJuntas = True

#End If

End Function

'genera los otros docs de tipo cheque que quedaron aprobados del año anterior
Public Function TraerOtrosDocsAprobados(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False) As Boolean
#If DATACON = 1 Then       'Access
   
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String
   Dim ConnStr As String
   Dim Condicion As String
   Dim OtrosDocs As Long
   Dim WhCheques As String
   
   RutMdb = Rut & ".mdb"
   
   TraerOtrosDocsAprobados = False
   
   If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
      If Msg Then
         MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar documentos pendientes en forma automática.", vbExclamation + vbOKOnly
      End If
      Exit Function
   End If
   
   'cerramos el año actual y abrimos el año anterior
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano - 1)
   Call LinkMdbAdm
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible traer cheques aprobados.", vbExclamation + vbOKOnly
      End If
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
      
      Exit Function
   End If
   
   If MsgBox1("Atención:" & vbNewLine & vbNewLine & "Sólo se traerán los cheques del año anterior que están en estado Aprobado y que no hayan sido importados con anterioridad." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
      
      Exit Function
   End If
      
   WhCheques = GenLike(DbMain, "CHEQUE", "TipoDocs.Nombre")
      
   'veamos si quedan cheques del año anterior en Estado Pendiente, si es así, mandamos mensaje
   Q1 = "SELECT IdDoc FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TIpoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Estado = " & ED_PENDIENTE & " AND Documento.TipoLib = " & LIB_OTROS & " AND " & WhCheques
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If Msg Then
         MsgBox1 "Hay cheques del año anterior que se encuentran en estado Pendiente. " & vbNewLine & vbNewLine & "Sólo se traerán los cheques que están en estado Aprobado.", vbExclamation + vbOKOnly
      End If
   End If
   
   Call CloseRs(Rs)
      
   'copiamos los IdDoc e IdMovDoc para despues re-vincular las tablas
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TIpoDocs.TipoDoc "
   Q1 = Q1 & " SET OldIdDocTmp = IdDoc, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND Estado = " & ED_APROBADO & " AND Documento.TipoLib = " & LIB_OTROS & " AND " & WhCheques
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
  
   
   'linkeamos la tabla de documentos del año actual para agregar los documentos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Documento", "DocumentoNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de entidades del año actual para agregar nuevas entidades, si las hay
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "Entidades", "EntidadesNew", , , gEmpresa.ConnStr)
      
   'primero traemos las entidades nuevas (dado los índices definidos en la tabla Entidades, sólo se insertarán los Ruts y Códigos nuevos (son únicos))
   '28 Sep 2006
   Q1 = "INSERT INTO EntidadesNew SELECT Entidades.* FROM Entidades "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa

   Call ExecSQL(DbMain, Q1)
      
   Where = " WHERE FExported < 0"
      
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan documentos repetidos
   Q1 = "UPDATE Documento INNER JOIN DocumentoNew ON Documento.idDoc = DocumentoNew.OldIdDoc "
   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Documento.IdEmpresa AND DocumentoNew.Ano - 1 = Documento.Ano "
   Q1 = Q1 & " SET Documento.FExported=" & CLng(Int(Now))
   Q1 = Q1 & " WHERE Documento.FExported < 0"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'Agregamos docs marcados para exportar
   Q1 = "INSERT INTO DocumentoNew"
   Q1 = Q1 & " SELECT IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " 0  As TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   'Q1 = Q1 & " ORDER BY Documento.FEmision"
   Call ExecSQL(DbMain, Q1)
   
   'NO insertamos los MovDocumento, porque los Otros documentos no tienen movimientos
         
   'actualizamos el Id de la entidad en los documentos calzando por RUT
   Q1 = "UPDATE (DocumentoNew"
   Q1 = Q1 & " INNER JOIN Entidades ON DocumentoNew.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = Entidades.IdEmpresa  )"
   Q1 = Q1 & " INNER JOIN EntidadesNew ON Entidades.Rut = EntidadesNew.Rut"
   Q1 = Q1 & " AND EntidadesNew.IdEmpresa = Entidades.IdEmpresa  "
   Q1 = Q1 & " SET DocumentoNew.IdEntidad = EntidadesNew.IdEntidad "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND DocumentoNew.IdEmpresa = " & IdEmpresa & " AND DocumentoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   
   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE DocumentoNew "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
      
   'Mensaje con cantidad
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      MsgBox1 "Se encontraron " & vFld(Rs("N")) & " cheques del año anterior en estado Aprobado.", vbInformation
   End If
   
   Call CloseRs(Rs)
      
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'soltamos las tablas
   Q1 = "DROP TABLE DocumentoNew"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "DROP TABLE EntidadesNew"
   Call ExecSQL(DbMain, Q1)
   
   'cerramos el año anterior y abrimos el año actual
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano)
   
   TraerOtrosDocsAprobados = True

#End If

End Function
'genera los activos fijos que tienen depreciación residual para el año siguiente
Public Function GenActFijoResidual(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String
   Dim ConnStr As String
   Dim NImported As Long
   
   If Not gEmprSeparadas Then
      GenActFijoResidual = GenActFijoResidualEmpJuntas(IdEmpresa, Rut, Ano, Msg, ClearFExported)
      Exit Function
   End If
   
#If DATACON = 1 Then       'Access
   
   RutMdb = Rut & ".mdb"
   
   GenActFijoResidual = False
   
   'existe año anterior?
   If Not ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) Then
      If Msg Then
         MsgBox1 "No se encontró la base de datos del año anterior. No es posible obtener los activos fijos que aún tienen vida útil residual en forma automática.", vbExclamation + vbOKOnly
      End If
      Exit Function
   End If
   
   If Not ClearFExported Then
      If MsgBox1("¿Desea traer los activos fijos que aún tienen vida útil residual desde el año anterior?", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
   
   'cerramos el año actual y abrimos el año anterior
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano - 1)
   
   Call LinkMdbAdm
   
   'corrige base del año anterior, por si las moscas
   Call CorrigeBase
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible obtener la información de los activos fijos que aún tienen vida útil residual.", vbExclamation + vbOKOnly
      End If
      
      'cerramos el año anterior y abrimos el año actual
      Call CloseDb(DbMain)
      
      Call OpenDbEmp(Rut, Ano)
      
      Exit Function
   End If
   
   
   'recalculamos los saldos finales de los de los activos fijos del año anterior
   If Not RecalcDepResidual(Ano - 1, Msg) Then
      Exit Function
   End If
   
   'limpiamos fecha de exportación en año anterior, si corresponde
   If ClearFExported Then       'FCA 18 Oct 2011: se desmarcan todos para que se puedan importar todos, salvo los que ya están (eso se controla más abajo)
   
      'Esta opción es sólo para el caso en que esté volviendo a generar el año
      'y haya borrado el archivo MDB del nuevo año por debajo.
      'Eso se detecta en ReadEmpresa
   
      Q1 = "UPDATE MovActivoFijo SET FExported = 0 WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1 'esto en el año anterior
      Call ExecSQL(DbMain, Q1)
   End If
      
   'marcamos los que vamos a exportar con -1
   'copiamos IdActFijo para después revincular
   Q1 = "UPDATE MovActivoFijo SET IdActFijoOldTmp = IdActFijo, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND ( VidaUtilResidual > 0 OR NoDepreciable <> 0 OR ValorLibro = 1 )"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   Where = " WHERE FExported < 0"
   
   
   'asignamos el IdOldTmp de las tablas adicionales al activo fijo
   Q1 = "UPDATE ActFijoFicha SET IdFichaOldTmp = IdFicha "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE ActFijoCompsFicha SET IdCompFichaOldTmp = IdCompFicha "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   
   
   'linkeamos la tabla de MovActivoFijo del año actual para agregar los activos fijos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "MovActivoFijo", "MovActivoFijoNew", , , gEmpresa.ConnStr)
   
   'linkeamos la tabla de ActFijoFicha del año actual para agregar los activos fijos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "ActFijoFicha", "ActFijoFichaNew", , , gEmpresa.ConnStr)
         
   'linkeamos la tabla de ActFijoCompsFicha del año actual para agregar los activos fijos, a partir del año anterior
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano & "\" & RutMdb, "ActFijoCompsFicha", "ActFijoCompsFichaNew", , , gEmpresa.ConnStr)
         
         
   'marcamos para exportar aquellos activos fijos que ya fueron exportados pero fueron eliminados en el nuevo año
   Q1 = "UPDATE MovActivoFijo LEFT JOIN MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
   Q1 = Q1 & " SET MovActivoFijo.FExported = -1"
   Q1 = Q1 & " WHERE MovActivoFijo.FExported >  0 AND MovActivoFijoNew.IdActFijo IS NULL AND ( MovActivoFijo.VidaUtilResidual > 0 OR MovActivoFijo.NoDepreciable <> 0 OR MovActivoFijo.ValorLibro = 1 )"
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1
   
   
   'prueba
'   Q1 = "UPDATE MovActivoFijo LEFT JOIN MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
'   Q1 = Q1 & " SET MovActivoFijo.FExported = -1"
'   Q1 = Q1 & " WHERE MovActivoFijo.FExported <  0  AND ( MovActivoFijo.VidaUtilResidual > 0 OR MovActivoFijo.NoDepreciable <> 0 OR MovActivoFijo.ValorLibro = 1 )"
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1
   
   
   Call ExecSQL(DbMain, Q1)
          
         
   'actualizamos los traidos con anterioridad: 1 MovActivoFijo
   Q1 = "UPDATE MovActivoFijoNew INNER JOIN MovActivoFijo ON MovActivoFijoNew.IdActFijoOld = MovActivoFijo.IdActFijo"
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
   Q1 = Q1 & " SET "
   Q1 = Q1 & "  MovActivoFijoNew.TipoDepHist = MovActivoFijo.TipoDep"
   Q1 = Q1 & ", MovActivoFijoNew.TipoDep = MovActivoFijo.TipoDep"
   Q1 = Q1 & ", MovActivoFijoNew.TipoDepLey21210Hist = MovActivoFijo.TipoDepLey21210"
'   Q1 = Q1 & ", MovActivoFijoNew.TipoDepLey21210 = iif( MovActivoFijo.TipoDepLey21210 = " & DEP_LEY21210_ARAUCANIA & ", 0, MovActivoFijo.TipoDepLey21210) "
   Q1 = Q1 & ", MovActivoFijoNew.TipoDepLey21210 = MovActivoFijo.TipoDepLey21210 "
   Q1 = Q1 & ", MovActivoFijoNew.DepLey21256Hist = MovActivoFijo.DepLey21256"
   
   Q1 = Q1 & ", MovActivoFijoNew.DepNormalHist = iif( MovActivoFijo.TipoDep = " & DEP_NORMAL & ",MovActivoFijo.DepNormalHist + MovActivoFijo.DepNormal, 0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepAceleradaHist = iif( MovActivoFijo.TipoDep = " & DEP_ACELERADA & ", MovActivoFijo.DepAceleradaHist + MovActivoFijo.DepAcelerada,0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepInstantHist = iif( MovActivoFijo.TipoDep = " & DEP_INSTANTANEA & ", MovActivoFijo.DepInstantHist + MovActivoFijo.DepInstant,0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepDecimaParteHist = iif( MovActivoFijo.TipoDep = " & DEP_DECIMAPARTE & ", MovActivoFijo.DepDecimaParteHist + MovActivoFijo.DepDecimaParte,0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepDecimaParte2Hist = iif( MovActivoFijo.TipoDep = " & DEP_DECIMAPARTE2 & ", MovActivoFijo.DepDecimaParte2Hist + MovActivoFijo.DepDecimaParte2,0)"
   
   Q1 = Q1 & ", MovActivoFijoNew.Neto = MovActivoFijo.Neto"
   Q1 = Q1 & ", MovActivoFijoNew.DepAcumHist = MovActivoFijo.DepAcumFinal"
   
   Q1 = Q1 & ", MovActivoFijoNew.DepNormal = iif( MovActivoFijo.TipoDep = " & DEP_NORMAL & ",iif( MovActivoFijo.VidaUtilResidual >= 12, 12, MovActivoFijo.VidaUtilResidual), 0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepAcelerada = iif( MovActivoFijo.TipoDep = " & DEP_ACELERADA & ", iif( MovActivoFijo.VidaUtilResidual >= 12, 12, MovActivoFijo.VidaUtilResidual), 0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepInstant = iif( MovActivoFijo.TipoDep = " & DEP_INSTANTANEA & ", iif( MovActivoFijo.VidaUtilResidual >= 12, 12, MovActivoFijo.VidaUtilResidual), 0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepDecimaParte = iif( MovActivoFijo.TipoDep = " & DEP_DECIMAPARTE & ", iif( MovActivoFijo.VidaUtilResidual >= 12, 12, MovActivoFijo.VidaUtilResidual), 0)"
   Q1 = Q1 & ", MovActivoFijoNew.DepDecimaParte2 = iif( MovActivoFijo.TipoDep = " & DEP_DECIMAPARTE2 & ", iif( MovActivoFijo.VidaUtilResidual >= 12, 12, MovActivoFijo.VidaUtilResidual), 0)"
   
   Q1 = Q1 & ", MovActivoFijoNew.TotalmenteDepreciado = iif(MovActivoFijo.ValorLibro = 1, 1, MovActivoFijo.TotalmenteDepreciado) "
   Q1 = Q1 & ", MovActivoFijoNew.ValReajustadoNetoAnt = MovActivoFijo.ValReajustadoNeto"
   Q1 = Q1 & ", MovActivoFijoNew.Cred4PorcAnoInit = iif(MovActivoFijo.Cred4Porc <> 0, MovActivoFijo.Cred4Porc, MovActivoFijo.Cred4PorcAnoInit)"    '17/08/12 Fca: Se agrega Cred4PorcAnoInit para almacenar si aplicó Cred33 en el primer año .

   Q1 = Q1 & " WHERE MovActivoFijoNew.IdEmpresa = " & IdEmpresa & " AND MovActivoFijoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   
   'actualizamos los traidos con anterioridad: 2 ActFijoFicha   EN PRINCIPIO NO ES NECESARIO
'   Q1 = "UPDATE ActFijoFichaNew INNER JOIN ActFijoFicha ON ActFijoFichaNew.IdFichaOld = ActFijoFicha.IdFicha"
'   Q1 = Q1 & " SET "
'   Q1 = Q1 & "   ActFijoFichaNew.PrecioFactura = ActFijoFicha.PrecioFactura"
'   Q1 = Q1 & ",  ActFijoFichaNew.DerechosIntern = ActFijoFicha.DerechosIntern"
'   Q1 = Q1 & ",  ActFijoFichaNew.Transporte = ActFijoFicha.Transporte"
'   Q1 = Q1 & ",  ActFijoFichaNew.ObrasAdapt = ActFijoFicha.ObrasAdapt"
'   Q1 = Q1 & ",  ActFijoFichaNew.PrecioAdquis = ActFijoFicha.ObrasAdapt"
'   Q1 = Q1 & ",  ActFijoFichaNew.IVARecuperable = ActFijoFicha.IVARecuperable"
'   Q1 = Q1 & ",  ActFijoFichaNew.FormacionPers = ActFijoFicha.FormacionPers"
'   Q1 = Q1 & ",  ActFijoFichaNew.ObrasReubic = ActFijoFicha.ObrasReubic"
'   Q1 = Q1 & ",  ActFijoFichaNew.TotalGastos = ActFijoFicha.TotalGastos"
'   Q1 = Q1 & ",  ActFijoFichaNew.FechaIncorporacion = ActFijoFicha.FechaIncorporacion"
'   Q1 = Q1 & ",  ActFijoFichaNew.FechaDisponible = ActFijoFicha.FechaDisponible"
'   Q1 = Q1 & ",  ActFijoFichaNew.AdquiOtrosConceptos = ActFijoFicha.AdquiOtrosConceptos"
'   Q1 = Q1 & ",  ActFijoFichaNew.GastoOtrosConceptos = ActFijoFicha.GastoOtrosConceptos"
'   Q1 = Q1 & ",  ActFijoFichaNew.SinDetComps = ActFijoFicha.SinDetComps"
'
'   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los traidos con anterioridad: 3 ActFijoCompsFicha
   Q1 = "UPDATE ActFijoCompsFichaNew INNER JOIN ActFijoCompsFicha ON ActFijoCompsFichaNew.IdCompFichaOld = ActFijoCompsFicha.IdCompFicha"
   Q1 = Q1 & " AND ActFijoCompsFicha.IdEmpresa = ActFijoCompsFichaNew.IdEmpresa AND ActFijoCompsFicha.Ano = ActFijoCompsFichaNew.Ano-1"
   Q1 = Q1 & " SET "
'   Q1 = Q1 & "   ActFijoCompsFichaNew.PjeDivComp = ActFijoCompsFicha.PjeDivComp"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.ValorCompra = ActFijoCompsFicha.ValorCompra"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.ValorResidual = ActFijoCompsFicha.ValorResidual"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.PjeAmortizacion = ActFijoCompsFicha.PjeAmortizacion"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.VidaUtil = ActFijoCompsFicha.VidaUtil"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.CostosAdicionales = ActFijoCompsFicha.CostosAdicionales"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.TasaDesc = ActFijoCompsFicha.TasaDesc"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.CostoDesmant = ActFijoCompsFicha.CostoDesmant"
'   Q1 = Q1 & ",  ActFijoCompsFichaNew.ValActCostoDesmant = ActFijoCompsFicha.ValActCostoDesmant"
   Q1 = Q1 & "  ActFijoCompsFichaNew.ValorBien = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.ValorBien * ActFijoCompsFicha.Factor)"
'   Q1 = Q1 & ", ActFijoCompsFichaNew.DepAcumuladaAnoAnt = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.DepPeriodo, (ActFijoCompsFicha.DepPeriodo + ActFijoCompsFicha.DepAcum)*ActFijoCompsFicha.Factor)"  'Joshua Catrin 1/02/2019
   Q1 = Q1 & ", ActFijoCompsFichaNew.DepAcumuladaAnoAnt = iif( ActFijoCompsFicha.Revalorizacion = 0, ActFijoCompsFicha.DepPeriodo + + ActFijoCompsFicha.DepAcum, (ActFijoCompsFicha.DepPeriodo + ActFijoCompsFicha.DepAcum)*ActFijoCompsFicha.Factor)"
'   Q1 = Q1 & ", ActFijoCompsFichaNew.VidaUtilYaDep = ActFijoCompsFicha.VidaUtilDep"
   Q1 = Q1 & ", ActFijoCompsFichaNew.VidaUtilYaDep = ActFijoCompsFicha.VidaUtilDep + ActFijoCompsFicha.VidaUtilYaDep "
   Q1 = Q1 & ", ActFijoCompsFichaNew.ReservaAcumAnt = ActFijoCompsFicha.ReservaAcum"
   Q1 = Q1 & " WHERE ActFijoCompsFichaNew.IdEmpresa = " & IdEmpresa & " AND ActFijoCompsFichaNew.Ano = " & Ano
   
   Call ExecSQL(DbMain, Q1)
      
   
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan activos fijos duplicados
   Q1 = "UPDATE MovActivoFijo INNER JOIN MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
   Q1 = Q1 & " SET MovActivoFijo.FExported=" & CLng(Int(Now))
   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0"
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)

   
   'Agregamos activos fijos marcados para exportar
   
   '30/09/10 Fca: Cred4Porc = 0 y ValCred33 = -1, dado que este beneficio sólo es válido para el primer año de compra del bien (solicitado por Victor Morales)
   '17/08/12 Fca: Se agrega Cred4PorcAnoInit para almacenar si aplicó Cred33 en el primer año.
   
   'En el reporte de ActFijo se recalcula el ValorLibro de este año
      
   Q1 = "INSERT INTO MovActivoFijoNew "
   Q1 = Q1 & " ( IdDoc, IdComp, IdMovComp, TipoMovAF, Fecha, Cantidad, Descrip, Neto, IVA, Cred4Porc, DepNormal, DepAcelerada, IdCuenta, DepNormalHist, DepAceleradaHist, NetoVenta, IVAVenta, FechaVentaBaja, TipoDep, TipoDepHist, DepAcumHist, VidaUtil, DepAcumFinal, VidaUtilResidual, FExported, FechaUtilizacion, NoDepreciable, ValCred33, ValReajustadoNeto, IdActFijoOld, TotalmenteDepreciado, ValorLibro, Cred4PorcAnoInit, DepInstant, DepDecimaParte, DepInstantHist, DepDecimaParteHist, VidaUtilAnos, TipoDepLey21210, DepDecimaParte2, DepDecimaParte2Hist, PatenteRol, NombreProy, FechaProy, TipoDepLey21210Hist, DepLey21256Hist, IdEmpresa, Ano)"
   Q1 = Q1 & " SELECT 0, 0, 0, MovActivoFijo.TipoMovAF, MovActivoFijo.Fecha, MovActivoFijo.Cantidad, MovActivoFijo.Descrip, MovActivoFijo.Neto, MovActivoFijo.IVA, 0, MovActivoFijo.DepNormal, MovActivoFijo.DepAcelerada, MovActivoFijo.IdCuenta, MovActivoFijo.DepNormalHist, MovActivoFijo.DepAceleradaHist, MovActivoFijo.NetoVenta, MovActivoFijo.IVAVenta, MovActivoFijo.FechaVentaBaja, MovActivoFijo.TipoDep, MovActivoFijo.TipoDepHist, MovActivoFijo.DepAcumHist, MovActivoFijo.VidaUtil, MovActivoFijo.DepAcumFinal, MovActivoFijo.VidaUtilResidual, MovActivoFijo.FExported, MovActivoFijo.FechaUtilizacion, MovActivoFijo.NoDepreciable, -1, MovActivoFijo.ValReajustadoNeto, MovActivoFijo.IdActFijoOldTmp, MovActivoFijo.TotalmenteDepreciado, MovActivoFijo.ValorLibro, iif(Cred4Porc <> 0, Cred4Porc, Cred4PorcAnoInit), MovActivoFijo.DepInstant, MovActivoFijo.DepDecimaParte, MovActivoFijo.DepInstantHist, MovActivoFijo.DepDecimaParteHist, MovActivoFijo.VidaUtilAnos "
   Q1 = Q1 & ", MovActivoFijo.TipoDepLey21210, MovActivoFijo.DepDecimaParte2, MovActivoFijo.DepDecimaParte2Hist, MovActivoFijo.PatenteRol, MovActivoFijo.NombreProy, MovActivoFijo.FechaProy, MovActivoFijo.TipoDepLey21210Hist, MovActivoFijo.DepLey21256Hist "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijo.Fecha "
   Call ExecSQL(DbMain, Q1)
   
   
   Q1 = "INSERT INTO ActFijoFichaNew "
   Q1 = Q1 & " ( IdActFijo, IdGrupo, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, SinDetComps, IdFichaOld, IdEmpresa, Ano  )"
   Q1 = Q1 & " SELECT MovActivoFijoNew.IdActFijo, ActFijoFicha.IdGrupo, ActFijoFicha.PrecioFactura, ActFijoFicha.DerechosIntern, ActFijoFicha.Transporte, ActFijoFicha.ObrasAdapt, ActFijoFicha.PrecioAdquis, ActFijoFicha.IVARecuperable, ActFijoFicha.FormacionPers, ActFijoFicha.ObrasReubic, ActFijoFicha.TotalGastos, ActFijoFicha.FechaIncorporacion, ActFijoFicha.FechaDisponible, ActFijoFicha.AdquiOtrosConceptos, ActFijoFicha.GastoOtrosConceptos, ActFijoFicha.SinDetComps, ActFijoFicha.IdFichaOldTmp "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ActFijoFicha INNER JOIN MovActivoFijoNew ON ActFijoFicha.IdActFijo =  MovActivoFijoNew.IdActFijoOld "
   Q1 = Q1 & "  AND ActFijoFicha.IdEmpresa = MovActivoFijoNew.IdEmpresa AND ActFijoFicha.Ano = MovActivoFijoNew.Ano - 1 "
   Q1 = Q1 & " WHERE MovActivoFijoNew.FExported < 0 "
   Q1 = Q1 & " AND ActFijoFicha.IdEmpresa = " & IdEmpresa & " AND ActFijoFicha.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijoNew.IdActFijo"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO ActFijoCompsFichaNew "
   Q1 = Q1 & " ( IdActFijo, IdGrupo, IdComp, PjeDivComp, ValorCompra, ValorResidual, PjeAmortizacion, VidaUtil, CostosAdicionales, TasaDesc, CostoDesmant, ValActCostoDesmant, ValorBien, ValorRazonable_31_12, NoExisteValRazonable, OtrasDiferencias, DepAcum, VidaUtilDep, ReservaAcum, DepAcumuladaAnoAnt, VidaUtilYaDep, ReservaAcumAnt,  IdCompFichaOld, DepPeriodo, Factor, Revalorizacion, IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT MovActivoFijoNew.IdActFijo,  ActFijoCompsFicha.IdGrupo, ActFijoCompsFicha.IdComp, ActFijoCompsFicha.PjeDivComp, ActFijoCompsFicha.ValorCompra, ActFijoCompsFicha.ValorResidual, ActFijoCompsFicha.PjeAmortizacion, ActFijoCompsFicha.VidaUtil, ActFijoCompsFicha.CostosAdicionales, ActFijoCompsFicha.TasaDesc, ActFijoCompsFicha.CostoDesmant, ActFijoCompsFicha.ValActCostoDesmant, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.ValorRazonable_31_12, ActFijoCompsFicha.NoExisteValRazonable, ActFijoCompsFicha.OtrasDiferencias, ActFijoCompsFicha.DepAcum, ActFijoCompsFicha.VidaUtilDep, ActFijoCompsFicha.ReservaAcum, ActFijoCompsFicha.DepAcumuladaAnoAnt, ActFijoCompsFicha.VidaUtilYaDep, ActFijoCompsFicha.ReservaAcumAnt, ActFijoCompsFicha.IdCompFichaOldTmp, ActFijoCompsFicha.DepPeriodo, ActFijoCompsFicha.Factor, ActFijoCompsFicha.Revalorizacion "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ActFijoCompsFicha INNER JOIN MovActivoFijoNew ON ActFijoCompsFicha.IdActFijo =  MovActivoFijoNew.IdActFijoOld"
   Q1 = Q1 & "  AND ActFijoCompsFicha.IdEmpresa = MovActivoFijoNew.IdEmpresa AND ActFijoCompsFicha.Ano = MovActivoFijoNew.Ano - 1 "
   Q1 = Q1 & " WHERE MovActivoFijoNew.FExported < 0 "
   Q1 = Q1 & " AND ActFijoCompsFicha.IdEmpresa = " & IdEmpresa & " AND ActFijoCompsFicha.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijoNew.IdActFijo"
   Call ExecSQL(DbMain, Q1)
   
     
   'Obtenemos cantidad de registros importados
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      NImported = vFld(Rs("N"))
   End If
   
   Call CloseRs(Rs)
     
   'actualizamos campos de ActFijoCompsFicha
   Q1 = "UPDATE ActFijoCompsFichaNew INNER JOIN MovActivoFijoNew ON ActFijoCompsFichaNew.IdActFijo =  MovActivoFijoNew.IdActFijo "
   Q1 = Q1 & " AND ActFijoCompsFichaNew.IdEmpresa = MovActivoFijoNew.IdEmpresa AND ActFijoCompsFichaNew.Ano = MovActivoFijoNew.Ano"
   Q1 = Q1 & " SET "
   Q1 = Q1 & "  ActFijoCompsFichaNew.ValorBien = iif( ActFijoCompsFichaNew.NoExisteValRazonable <> 0, ActFijoCompsFichaNew.ValorBien, ActFijoCompsFichaNew.ValorBien * ActFijoCompsFichaNew.Factor)"
   Q1 = Q1 & ", ActFijoCompsFichaNew.DepAcumuladaAnoAnt = iif( ActFijoCompsFichaNew.NoExisteValRazonable <> 0, ActFijoCompsFichaNew.DepPeriodo, (ActFijoCompsFichaNew.DepPeriodo + ActFijoCompsFichaNew.DepAcum) * ActFijoCompsFichaNew.Factor )"
   Q1 = Q1 & ", ActFijoCompsFichaNew.VidaUtilYaDep = ActFijoCompsFichaNew.VidaUtilDep "
   Q1 = Q1 & ", ActFijoCompsFichaNew.ReservaAcumAnt = ActFijoCompsFichaNew.ReservaAcum "
   Q1 = Q1 & " WHERE MovActivoFijoNew.FExported < 0"
   Q1 = Q1 & " AND MovActivoFijoNew.IdEmpresa = " & IdEmpresa & " AND MovActivoFijoNew.Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
     
   'actualizamos campos de MovActiFijo
   Q1 = "UPDATE MovActivoFijoNew SET "
   Q1 = Q1 & "  IdDoc = 0 "
   Q1 = Q1 & ", IdComp = 0 "
   Q1 = Q1 & ", IdMovComp = 0 "
   Q1 = Q1 & ", TipoDepHist = TipoDep "
   Q1 = Q1 & ", TipoDepLey21210Hist = TipoDepLey21210 "
   Q1 = Q1 & ", DepLey21256Hist = DepLey21256 "
   
   Q1 = Q1 & ", DepNormalHist = iif( TipoDep = " & DEP_NORMAL & ", DepNormalHist + DepNormal, 0)"
   Q1 = Q1 & ", DepAceleradaHist = iif( TipoDep = " & DEP_ACELERADA & ", DepAceleradaHist + DepAcelerada, 0)"
   Q1 = Q1 & ", DepInstantHist = iif( TipoDep = " & DEP_INSTANTANEA & ", DepInstantHist + DepInstant, 0)"
   Q1 = Q1 & ", DepDecimaParteHist = iif( TipoDep = " & DEP_DECIMAPARTE & ", DepDecimaParteHist + DepDecimaParte, 0)"
   Q1 = Q1 & ", DepDecimaParte2Hist = iif( TipoDep = " & DEP_DECIMAPARTE2 & ", DepDecimaParte2Hist + DepDecimaParte2, 0)"
   
   Q1 = Q1 & ", ValReajustadoNetoAnt = ValReajustadoNeto"
   Q1 = Q1 & ", DepAcumHist = DepAcumFinal"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'limpiamos FExported en tabla nueva y actualizamos campos
   'actualizamos campos de depreciación en un query aparte para evitar que se pisen los cambios
   
   Q1 = "UPDATE MovActivoFijoNew SET "
   Q1 = Q1 & "  FExported = 0"
   Q1 = Q1 & ", FImported = " & CLng(Int(Now))
   Q1 = Q1 & ", TipoDep = iif(TipoDep = " & DEP_DECIMAPARTE & " or TipoDep = " & DEP_DECIMAPARTE2 & " or TipoDep = " & DEP_INSTANTANEA & ", " & DEP_NORMAL & ", TipoDep )"
   Q1 = Q1 & ", DepNormal = iif( TipoDep = " & DEP_NORMAL & " or TipoDep = " & DEP_DECIMAPARTE & " or TipoDep = " & DEP_DECIMAPARTE2 & " or TipoDep = " & DEP_INSTANTANEA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepAcelerada = iif( TipoDep = " & DEP_ACELERADA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepInstant = 0"
   Q1 = Q1 & ", DepDecimaParte = 0"
   Q1 = Q1 & ", DepDecimaParte2 = 0"
   Q1 = Q1 & ", TotalmenteDepreciado = iif(ValorLibro = 1, 1, TotalmenteDepreciado) "
   Q1 = Q1 & ", TipoDepLey21210 = 0 "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
            
   'Mensaje con cantidad
   MsgBox1 "Se importaron " & NImported & " Activos Fijos del año anterior, con vida útil residual o no depreciables.", vbInformation
   
   Call CloseRs(Rs)
   
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE MovActivoFijo "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   'soltamos las tablas
   Q1 = "DROP TABLE MovActivoFijoNew"
   Call ExecSQL(DbMain, Q1)
   Q1 = "DROP TABLE ActFijoFichaNew"
   Call ExecSQL(DbMain, Q1)
   Q1 = "DROP TABLE ActFijoCompsFichaNew"
   Call ExecSQL(DbMain, Q1)
   
   'cerramos el año anterior y abrimos el año actual
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(Rut, Ano)
   
   GenActFijoResidual = True

#End If

End Function
'genera los activos fijos que tienen depreciación residual para el año siguiente, en bases juntas
Public Function GenActFijoResidualEmpJuntas(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal Msg As Boolean = False, Optional ByVal ClearFExported As Boolean = False) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim Where As String
   Dim ConnStr As String
   Dim NImported As Long
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
   GenActFijoResidualEmpJuntas = False
   
   If gEmprSeparadas Then
      Exit Function
   End If
   
   If gEmpresa.TieneAnoAntAccess Then  'los Activos Fijos Residuales ya fueron generados al crear el nuevo año desde Access
      GenActFijoResidualEmpJuntas = True
      Exit Function
   End If
   
   
   'vemos si el año anterior está cerrado
   Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      FCierre = vFld(Rs("FCierre"))
   Else
      MsgBox1 "No hay registro de año anterior. No es posible generar documentos pendientes.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Function
   End If
   
   Call CloseRs(Rs)
   
   If FCierre = 0 Then
      If Msg Then
         MsgBox1 "El año anterior aún no ha sido cerrado. No es posible obtener la información de los activos fijos que aún tienen vida útil residual.", vbExclamation + vbOKOnly
      End If
      
      Exit Function
   End If
   
   If Not ClearFExported Then
      If MsgBox1("¿Desea traer los activos fijos que aún tienen vida útil residual desde el año anterior?", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
   
   
   'recalculamos los saldos finales de los de los activos fijos del año anterior
   If Not RecalcDepResidual(Ano - 1, Msg) Then
      Exit Function
   End If
   
   'limpiamos fecha de exportación en año anterior, si corresponde
   If ClearFExported Then       'FCA 18 Oct 2011: se desmarcan todos para que se puedan importar todos, salvo los que ya están (eso se controla más abajo)
   
      'Esta opción es sólo para el caso en que esté volviendo a generar el año
      'y haya borrado el archivo MDB del nuevo año por debajo.
      'Eso se detecta en ReadEmpresa
   
      Q1 = "UPDATE MovActivoFijo SET FExported = 0 WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1 'esto en el año anterior
      Call ExecSQL(DbMain, Q1)
   End If
      
   'marcamos los que vamos a exportar con -1
   'copiamos IdActFijo para después revincular
   Q1 = "UPDATE MovActivoFijo SET IdActFijoOldTmp = IdActFijo, FExported = -1"
   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) AND ( VidaUtilResidual > 0 OR NoDepreciable <> 0 OR ValorLibro = 1 )"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   Where = " WHERE FExported < 0"
   
   
   'asignamos el IdOldTmp de las tablas adicionales al activo fijo
   Q1 = "UPDATE ActFijoFicha SET IdFichaOldTmp = IdFicha "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE ActFijoCompsFicha SET IdCompFichaOldTmp = IdCompFicha "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
   
         
   'marcamos para exportar aquellos activos fijos que ya fueron exportados pero fueron eliminados en el nuevo año
'   Q1 = "UPDATE MovActivoFijo LEFT JOIN MovActivoFijo as MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
'   Q1 = Q1 & " SET MovActivoFijo.FExported = -1 "
'   Q1 = Q1 & " WHERE MovActivoFijo.FExported > 0 AND MovActivoFijoNew.IdActFijo IS NULL "
'   Q1 = Q1 & " AND ( MovActivoFijo.VidaUtilResidual > 0 OR MovActivoFijo.NoDepreciable <> 0 OR MovActivoFijo.ValorLibro = 1 )"
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1
'   Call ExecSQL(DbMain, Q1)
   Tbl = " MovActivoFijo "
   sFrom = " MovActivoFijo LEFT JOIN MovActivoFijo as MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
   sFrom = sFrom & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
   sSet = " MovActivoFijo.FExported = -1 "
   sWhere = " WHERE MovActivoFijo.FExported > 0 AND MovActivoFijoNew.IdActFijo IS NULL "
   sWhere = sWhere & " AND ( MovActivoFijo.VidaUtilResidual > 0 OR MovActivoFijo.NoDepreciable <> 0 OR MovActivoFijo.ValorLibro = 1 )"
   sWhere = sWhere & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
          
   'actualizamos los traidos con anterioridad: 1 MovActivoFijo
'   Q1 = "UPDATE MovActivoFijo INNER JOIN MovActivoFijo as MovActivoFijoOld ON MovActivoFijo.IdActFijoOld = MovActivoFijoOld.IdActFijo"
'   Q1 = Q1 & " AND MovActivoFijoOld.IdEmpresa = MovActivoFijo.IdEmpresa AND MovActivoFijoOld.Ano = MovActivoFijo.Ano - 1"
'   Q1 = Q1 & " SET "
'   Q1 = Q1 & "  MovActivoFijo.TipoDepHist = MovActivoFijoOld.TipoDep"
'   Q1 = Q1 & ", MovActivoFijo.TipoDep = MovActivoFijoOld.TipoDep"
'
'   Q1 = Q1 & ", MovActivoFijo.DepNormalHist = iif( MovActivoFijoOld.TipoDep = " & DEP_NORMAL & ",MovActivoFijoOld.DepNormalHist + MovActivoFijoOld.DepNormal, 0)"
'   Q1 = Q1 & ", MovActivoFijo.DepAceleradaHist = iif( MovActivoFijoOld.TipoDep = " & DEP_ACELERADA & ", MovActivoFijoOld.DepAceleradaHist + MovActivoFijoOld.DepAcelerada,0)"
'   Q1 = Q1 & ", MovActivoFijo.DepInstantHist = iif( MovActivoFijoOld.TipoDep = " & DEP_INSTANTANEA & ", MovActivoFijoOld.DepInstantHist + MovActivoFijoOld.DepInstant,0)"
'   Q1 = Q1 & ", MovActivoFijo.DepDecimaParteHist = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE & ", MovActivoFijoOld.DepDecimaParteHist + MovActivoFijoOld.DepDecimaParte,0)"
'
'   Q1 = Q1 & ", MovActivoFijo.Neto = MovActivoFijoOld.Neto"
'   Q1 = Q1 & ", MovActivoFijo.DepAcumHist = MovActivoFijoOld.DepAcumFinal"
'
'   Q1 = Q1 & ", MovActivoFijo.DepNormal = iif( MovActivoFijoOld.TipoDep = " & DEP_NORMAL & ",iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
'   Q1 = Q1 & ", MovActivoFijo.DepAcelerada = iif( MovActivoFijoOld.TipoDep = " & DEP_ACELERADA & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
'   Q1 = Q1 & ", MovActivoFijo.DepInstant = iif( MovActivoFijoOld.TipoDep = " & DEP_INSTANTANEA & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
'   Q1 = Q1 & ", MovActivoFijo.DepDecimaParte = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
'
'
'   Q1 = Q1 & ", MovActivoFijo.TotalmenteDepreciado = iif(MovActivoFijoOld.ValorLibro = 1, 1, MovActivoFijoOld.TotalmenteDepreciado) "
'   Q1 = Q1 & ", MovActivoFijo.ValReajustadoNetoAnt = MovActivoFijoOld.ValReajustadoNeto"
'   Q1 = Q1 & ", MovActivoFijo.Cred4PorcAnoInit = iif(MovActivoFijoOld.Cred4Porc <> 0, MovActivoFijoOld.Cred4Porc, MovActivoFijoOld.Cred4PorcAnoInit)"    '17/08/12 Fca: Se agrega Cred4PorcAnoInit para almacenar si aplicó Cred33 en el primer año .
'
'   Q1 = Q1 & " WHERE MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   
   Tbl = " MovActivoFijo "
   sFrom = " MovActivoFijo INNER JOIN MovActivoFijo as MovActivoFijoOld ON MovActivoFijo.IdActFijoOld = MovActivoFijoOld.IdActFijo"
   sFrom = sFrom & " AND MovActivoFijoOld.IdEmpresa = MovActivoFijo.IdEmpresa AND MovActivoFijoOld.Ano = MovActivoFijo.Ano - 1"
   
   sSet = "  MovActivoFijo.TipoDepHist = MovActivoFijoOld.TipoDep"
   sSet = sSet & ", MovActivoFijo.TipoDep = MovActivoFijoOld.TipoDep"
   sSet = sSet & ", MovActivoFijo.TipoDepLey21210Hist = MovActivoFijoOld.TipoDepLey21210"
   sSet = sSet & ", MovActivoFijo.TipoDepLey21210 = MovActivoFijoOld.TipoDepLey21210 "
   sSet = sSet & ", MovActivoFijo.DepLey21256Hist = MovActivoFijoOld.DepLey21256"
   sSet = sSet & ", MovActivoFijo.DepLey21256 = MovActivoFijoOld.DepLey21256 "
  
   sSet = sSet & ", MovActivoFijo.DepNormalHist = iif( MovActivoFijoOld.TipoDep = " & DEP_NORMAL & ",MovActivoFijoOld.DepNormalHist + MovActivoFijoOld.DepNormal, 0)"
   sSet = sSet & ", MovActivoFijo.DepAceleradaHist = iif( MovActivoFijoOld.TipoDep = " & DEP_ACELERADA & ", MovActivoFijoOld.DepAceleradaHist + MovActivoFijoOld.DepAcelerada,0)"
   sSet = sSet & ", MovActivoFijo.DepInstantHist = iif( MovActivoFijoOld.TipoDep = " & DEP_INSTANTANEA & ", MovActivoFijoOld.DepInstantHist + MovActivoFijoOld.DepInstant,0)"
   sSet = sSet & ", MovActivoFijo.DepDecimaParteHist = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE & ", MovActivoFijoOld.DepDecimaParteHist + MovActivoFijoOld.DepDecimaParte,0)"
   sSet = sSet & ", MovActivoFijo.DepDecimaParte2Hist = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE2 & ", MovActivoFijoOld.DepDecimaParte2Hist + MovActivoFijoOld.DepDecimaParte2,0)"
   
   sSet = sSet & ", MovActivoFijo.Neto = MovActivoFijoOld.Neto"
   sSet = sSet & ", MovActivoFijo.DepAcumHist = MovActivoFijoOld.DepAcumFinal"
   
   sSet = sSet & ", MovActivoFijo.DepNormal = iif( MovActivoFijoOld.TipoDep = " & DEP_NORMAL & ",iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
   sSet = sSet & ", MovActivoFijo.DepAcelerada = iif( MovActivoFijoOld.TipoDep = " & DEP_ACELERADA & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
   sSet = sSet & ", MovActivoFijo.DepInstant = iif( MovActivoFijoOld.TipoDep = " & DEP_INSTANTANEA & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
   sSet = sSet & ", MovActivoFijo.DepDecimaParte = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
   sSet = sSet & ", MovActivoFijo.DepDecimaParte2 = iif( MovActivoFijoOld.TipoDep = " & DEP_DECIMAPARTE2 & ", iif( MovActivoFijoOld.VidaUtilResidual >= 12, 12, MovActivoFijoOld.VidaUtilResidual), 0)"
   
   
   sSet = sSet & ", MovActivoFijo.TotalmenteDepreciado = iif(MovActivoFijoOld.ValorLibro = 1, 1, MovActivoFijoOld.TotalmenteDepreciado) "
   sSet = sSet & ", MovActivoFijo.ValReajustadoNetoAnt = MovActivoFijoOld.ValReajustadoNeto"
   sSet = sSet & ", MovActivoFijo.Cred4PorcAnoInit = iif(MovActivoFijoOld.Cred4Porc <> 0, MovActivoFijoOld.Cred4Porc, MovActivoFijoOld.Cred4PorcAnoInit)"    '17/08/12 Fca: Se agrega Cred4PorcAnoInit para almacenar si aplicó Cred33 en el primer año .

   sWhere = " WHERE MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'actualizamos los traidos con anterioridad: 2 ActFijoFicha   EN PRINCIPIO NO ES NECESARIO
'   Q1 = "UPDATE ActFijoFichaNew INNER JOIN ActFijoFicha ON ActFijoFichaNew.IdFichaOld = ActFijoFicha.IdFicha"
'   Q1 = Q1 & " SET "
'   Q1 = Q1 & "   ActFijoFichaNew.PrecioFactura = ActFijoFicha.PrecioFactura"
'   Q1 = Q1 & ",  ActFijoFichaNew.DerechosIntern = ActFijoFicha.DerechosIntern"
'   Q1 = Q1 & ",  ActFijoFichaNew.Transporte = ActFijoFicha.Transporte"
'   Q1 = Q1 & ",  ActFijoFichaNew.ObrasAdapt = ActFijoFicha.ObrasAdapt"
'   Q1 = Q1 & ",  ActFijoFichaNew.PrecioAdquis = ActFijoFicha.ObrasAdapt"
'   Q1 = Q1 & ",  ActFijoFichaNew.IVARecuperable = ActFijoFicha.IVARecuperable"
'   Q1 = Q1 & ",  ActFijoFichaNew.FormacionPers = ActFijoFicha.FormacionPers"
'   Q1 = Q1 & ",  ActFijoFichaNew.ObrasReubic = ActFijoFicha.ObrasReubic"
'   Q1 = Q1 & ",  ActFijoFichaNew.TotalGastos = ActFijoFicha.TotalGastos"
'   Q1 = Q1 & ",  ActFijoFichaNew.FechaIncorporacion = ActFijoFicha.FechaIncorporacion"
'   Q1 = Q1 & ",  ActFijoFichaNew.FechaDisponible = ActFijoFicha.FechaDisponible"
'   Q1 = Q1 & ",  ActFijoFichaNew.AdquiOtrosConceptos = ActFijoFicha.AdquiOtrosConceptos"
'   Q1 = Q1 & ",  ActFijoFichaNew.GastoOtrosConceptos = ActFijoFicha.GastoOtrosConceptos"
'   Q1 = Q1 & ",  ActFijoFichaNew.SinDetComps = ActFijoFicha.SinDetComps"
'
'   Call ExecSQL(DbMain, Q1)
   
   'actualizamos los traidos con anterioridad: 3 ActFijoCompsFicha
'   Q1 = "UPDATE ActFijoCompsFicha INNER JOIN ActFijoCompsFicha as  ActFijoCompsFichaOld ON ActFijoCompsFicha.IdCompFichaOld = ActFijoCompsFichaOld.IdCompFicha"
'   Q1 = Q1 & " AND ActFijoCompsFichaOld.IdEmpresa = ActFijoCompsFicha.IdEmpresa AND ActFijoCompsFichaOld.Ano = ActFijoCompsFicha.Ano-1"
'   Q1 = Q1 & " SET "
''   Q1 = Q1 & "   ActFijoCompsFicha.PjeDivComp = ActFijoCompsFichaOld.PjeDivComp"
''   Q1 = Q1 & ",  ActFijoCompsFicha.ValorCompra = ActFijoCompsFichaOld.ValorCompra"
''   Q1 = Q1 & ",  ActFijoCompsFicha.ValorResidual = ActFijoCompsFichaOld.ValorResidual"
''   Q1 = Q1 & ",  ActFijoCompsFicha.PjeAmortizacion = ActFijoCompsFichaOld.PjeAmortizacion"
''   Q1 = Q1 & ",  ActFijoCompsFicha.VidaUtil = ActFijoCompsFichaOld.VidaUtil"
''   Q1 = Q1 & ",  ActFijoCompsFicha.CostosAdicionales = ActFijoCompsFichaOld.CostosAdicionales"
''   Q1 = Q1 & ",  ActFijoCompsFicha.TasaDesc = ActFijoCompsFichaOld.TasaDesc"
''   Q1 = Q1 & ",  ActFijoCompsFicha.CostoDesmant = ActFijoCompsFichaOld.CostoDesmant"
''   Q1 = Q1 & ",  ActFijoCompsFicha.ValActCostoDesmant = ActFijoCompsFichaOld.ValActCostoDesmant"
'   Q1 = Q1 & "  ActFijoCompsFicha.ValorBien = iif( ActFijoCompsFichaOld.NoExisteValRazonable <> 0, ActFijoCompsFichaOld.ValorBien, ActFijoCompsFichaOld.ValorBien * ActFijoCompsFichaOld.Factor)"
'   Q1 = Q1 & ", ActFijoCompsFicha.DepAcumuladaAnoAnt = iif( ActFijoCompsFichaOld.NoExisteValRazonable <> 0, ActFijoCompsFichaOld.DepPeriodo, (ActFijoCompsFichaOld.DepPeriodo + ActFijoCompsFichaOld.DepAcum)*ActFijoCompsFichaOld.Factor)"
'   Q1 = Q1 & ", ActFijoCompsFicha.VidaUtilYaDep = ActFijoCompsFichaOld.VidaUtilDep"
'   Q1 = Q1 & ", ActFijoCompsFicha.ReservaAcumAnt = ActFijoCompsFichaOld.ReservaAcum"
'   Q1 = Q1 & " WHERE ActFijoCompsFicha.IdEmpresa = " & IdEmpresa & " AND ActFijoCompsFicha.Ano = " & Ano
'
'   Call ExecSQL(DbMain, Q1)
   Tbl = " ActFijoCompsFicha "
   sFrom = " ActFijoCompsFicha INNER JOIN ActFijoCompsFicha as  ActFijoCompsFichaOld ON ActFijoCompsFicha.IdCompFichaOld = ActFijoCompsFichaOld.IdCompFicha"
   sFrom = sFrom & " AND ActFijoCompsFichaOld.IdEmpresa = ActFijoCompsFicha.IdEmpresa AND ActFijoCompsFichaOld.Ano = ActFijoCompsFicha.Ano-1"
   sSet = " ActFijoCompsFicha.ValorBien = iif( ActFijoCompsFichaOld.NoExisteValRazonable <> 0, ActFijoCompsFichaOld.ValorBien, ActFijoCompsFichaOld.ValorBien * ActFijoCompsFichaOld.Factor)"
   sSet = sSet & ", ActFijoCompsFicha.DepAcumuladaAnoAnt = iif( ActFijoCompsFichaOld.NoExisteValRazonable <> 0, ActFijoCompsFichaOld.DepPeriodo, (ActFijoCompsFichaOld.DepPeriodo + ActFijoCompsFichaOld.DepAcum)*ActFijoCompsFichaOld.Factor)"
   sSet = sSet & ", ActFijoCompsFicha.VidaUtilYaDep = ActFijoCompsFichaOld.VidaUtilDep"
   sSet = sSet & ", ActFijoCompsFicha.ReservaAcumAnt = ActFijoCompsFichaOld.ReservaAcum"
   sWhere = " WHERE ActFijoCompsFicha.IdEmpresa = " & IdEmpresa & " AND ActFijoCompsFicha.Ano = " & Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
      
   
   'desmarcamos los que ya están (que no debería ser pero....) para que no aparezcan activos fijos duplicados
'   Q1 = "UPDATE MovActivoFijo INNER JOIN MovActivoFijo As MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
'   Q1 = Q1 & " SET MovActivoFijo.FExported=" & CLng(Int(Now))
'   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0"
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1
'   Call ExecSQL(DbMain, Q1)
   Tbl = " MovActivoFijo "
   sFrom = " MovActivoFijo INNER JOIN MovActivoFijo As MovActivoFijoNew ON MovActivoFijo.IdActFijo = MovActivoFijoNew.IdActFijoOld "
   sFrom = sFrom & " AND MovActivoFijo.IdEmpresa = MovActivoFijoNew.IdEmpresa AND MovActivoFijo.Ano = MovActivoFijoNew.Ano-1"
   sSet = " MovActivoFijo.FExported=" & CLng(Int(Now))
   sWhere = " WHERE MovActivoFijo.FExported < 0"
   sWhere = sWhere & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano - 1

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Agregamos activos fijos marcados para exportar
   
   '30/09/10 Fca: Cred4Porc = 0 y ValCred33 = -1, dado que este beneficio sólo es válido para el primer año de compra del bien (solicitado por Victor Morales)
   '17/08/12 Fca: Se agrega Cred4PorcAnoInit para almacenar si aplicó Cred33 en el primer año.
   
   'En el reporte de ActFijo se recalcula el ValorLibro de este año
      
   Q1 = "INSERT INTO MovActivoFijo "
   Q1 = Q1 & " ( IdDoc, IdComp, IdMovComp, TipoMovAF, Fecha, Cantidad, Descrip, Neto, IVA, Cred4Porc, DepNormal, DepAcelerada, IdCuenta, DepNormalHist, DepAceleradaHist, NetoVenta, IVAVenta, FechaVentaBaja, TipoDep, TipoDepHist, DepAcumHist, VidaUtil, DepAcumFinal, VidaUtilResidual, FExported, FechaUtilizacion, NoDepreciable, ValCred33, ValReajustadoNeto, IdActFijoOld, TotalmenteDepreciado, ValorLibro, Cred4PorcAnoInit, DepInstant, DepDecimaParte, DepInstantHist, DepDecimaParteHist, VidaUtilAnos, TipoDepLey21210, DepDecimaParte2, DepDecimaParte2Hist, PatenteRol, NombreProy, FechaProy, TipoDepLey21210Hist, DepLey21256Hist, IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT 0 as IdDoc, 0 As IdComp, 0 As IdMovComp, TipoMovAF, Fecha, Cantidad, Descrip, Neto, IVA, 0 As Cred4Porc, DepNormal, DepAcelerada, IdCuenta, DepNormalHist, DepAceleradaHist, NetoVenta, IVAVenta, FechaVentaBaja, TipoDep, TipoDepHist, DepAcumHist, VidaUtil, DepAcumFinal, VidaUtilResidual, FExported, FechaUtilizacion, NoDepreciable, -1 As ValCred33, ValReajustadoNeto, IdActFijoOldTmp As IdActFijoOld, TotalmenteDepreciado, ValorLibro, iif(MovActivoFijoOld.Cred4Porc <> 0, MovActivoFijoOld.Cred4Porc, MovActivoFijoOld.Cred4PorcAnoInit) As Cred4PorcAnoInit, DepInstant, DepDecimaParte, DepInstantHist, DepDecimaParteHist, VidaUtilAnos "
   Q1 = Q1 & ", TipoDepLey21210, DepDecimaParte2, DepDecimaParte2Hist, PatenteRol, NombreProy, FechaProy, TipoDepLey21210Hist, DepLey21256Hist"
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM MovActivoFijo as MovActivoFijoOld "
   Q1 = Q1 & " WHERE MovActivoFijoOld.FExported < 0 "
   Q1 = Q1 & " AND MovActivoFijoOld.IdEmpresa = " & IdEmpresa & " AND MovActivoFijoOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijoOld.Fecha "
   Call ExecSQL(DbMain, Q1)
   
   
   Q1 = "INSERT INTO ActFijoFicha "
   Q1 = Q1 & " ( IdActFijo, IdGrupo, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, SinDetComps, IdFichaOld, IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT MovActivoFijo.IdActFijo, IdGrupo, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, SinDetComps, IdFichaOldTmp As IdFichaOld "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ActFijoFicha as ActFijoFichaOld INNER JOIN MovActivoFijo ON ActFijoFichaOld.IdActFijo =  MovActivoFijo.IdActFijoOld "
   Q1 = Q1 & "  AND ActFijoFichaOld.IdEmpresa = MovActivoFijo.IdEmpresa AND ActFijoFichaOld.Ano = MovActivoFijo.Ano - 1 "
   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0 "
   Q1 = Q1 & " AND ActFijoFichaOld.IdEmpresa = " & IdEmpresa & " AND ActFijoFichaOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijo.IdActFijo"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO ActFijoCompsFicha "
   Q1 = Q1 & " ( IdActFijo, IdGrupo, IdComp, PjeDivComp, ValorCompra, ValorResidual, PjeAmortizacion, VidaUtil, CostosAdicionales, TasaDesc, CostoDesmant, ValActCostoDesmant, ValorBien, ValorRazonable_31_12, NoExisteValRazonable, OtrasDiferencias, DepAcum, VidaUtilDep, ReservaAcum, DepAcumuladaAnoAnt, VidaUtilYaDep, ReservaAcumAnt,  IdCompFichaOld, DepPeriodo, Factor, Revalorizacion, IdEmpresa, Ano  )"
   Q1 = Q1 & " SELECT MovActivoFijo.IdActFijo,  IdGrupo, ActFijoCompsFichaOld.IdComp, PjeDivComp, ValorCompra, ValorResidual, PjeAmortizacion, ActFijoCompsFichaOld.VidaUtil, CostosAdicionales, TasaDesc, CostoDesmant, ValActCostoDesmant, ValorBien, ValorRazonable_31_12, NoExisteValRazonable, OtrasDiferencias, DepAcum, VidaUtilDep, ReservaAcum, DepAcumuladaAnoAnt, VidaUtilYaDep, ReservaAcumAnt, IdCompFichaOldTmp As IdCompFichaOld, DepPeriodo, Factor, Revalorizacion "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM ActFijoCompsFicha as ActFijoCompsFichaOld INNER JOIN MovActivoFijo ON ActFijoCompsFichaOld.IdActFijo =  MovActivoFijo.IdActFijoOld"
   Q1 = Q1 & "  AND ActFijoCompsFichaOld.IdEmpresa = MovActivoFijo.IdEmpresa AND ActFijoCompsFichaOld.Ano = MovActivoFijo.Ano - 1 "
   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0 "
   Q1 = Q1 & " AND ActFijoCompsFichaOld.IdEmpresa = " & IdEmpresa & " AND ActFijoCompsFichaOld.Ano = " & Ano - 1
   Q1 = Q1 & " ORDER BY MovActivoFijo.IdActFijo"
   Call ExecSQL(DbMain, Q1)
   
     
   'Obtenemos cantidad de registros importados
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      NImported = vFld(Rs("N"))
   End If
   
   Call CloseRs(Rs)
     
   'actualizamos campos de ActFijoCompsFicha
'   Q1 = "UPDATE ActFijoCompsFicha INNER JOIN MovActivoFijo ON ActFijoCompsFicha.IdActFijo =  MovActivoFijo.IdActFijo "
'   Q1 = Q1 & " AND ActFijoCompsFicha.IdEmpresa = MovActivoFijo.IdEmpresa AND ActFijoCompsFicha.Ano = MovActivoFijo.Ano"
'   Q1 = Q1 & " SET "
'   Q1 = Q1 & "  ActFijoCompsFicha.ValorBien = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.ValorBien * ActFijoCompsFicha.Factor)"
'   Q1 = Q1 & ", ActFijoCompsFicha.DepAcumuladaAnoAnt = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.DepPeriodo, (ActFijoCompsFicha.DepPeriodo + ActFijoCompsFicha.DepAcum) * ActFijoCompsFicha.Factor )"
'   Q1 = Q1 & ", ActFijoCompsFicha.VidaUtilYaDep = ActFijoCompsFicha.VidaUtilDep "
'   Q1 = Q1 & ", ActFijoCompsFicha.ReservaAcumAnt = ActFijoCompsFicha.ReservaAcum "
'   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0"
'   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " ActFijoCompsFicha "
   sFrom = " ActFijoCompsFicha INNER JOIN MovActivoFijo ON ActFijoCompsFicha.IdActFijo =  MovActivoFijo.IdActFijo "
   sFrom = sFrom & " AND ActFijoCompsFicha.IdEmpresa = MovActivoFijo.IdEmpresa AND ActFijoCompsFicha.Ano = MovActivoFijo.Ano"
   sSet = "  ActFijoCompsFicha.ValorBien = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.ValorBien * ActFijoCompsFicha.Factor)"
   sSet = sSet & ", ActFijoCompsFicha.DepAcumuladaAnoAnt = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.DepPeriodo, (ActFijoCompsFicha.DepPeriodo + ActFijoCompsFicha.DepAcum) * ActFijoCompsFicha.Factor )"
   sSet = sSet & ", ActFijoCompsFicha.VidaUtilYaDep = ActFijoCompsFicha.VidaUtilDep "
   sSet = sSet & ", ActFijoCompsFicha.ReservaAcumAnt = ActFijoCompsFicha.ReservaAcum "
   sWhere = " WHERE MovActivoFijo.FExported < 0"
   sWhere = sWhere & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
     
   'actualizamos campos de MovActiFijo
   Q1 = "UPDATE MovActivoFijo SET "
   Q1 = Q1 & "  IdDoc = 0 "
   Q1 = Q1 & ", IdComp = 0 "
   Q1 = Q1 & ", IdMovComp = 0 "
   Q1 = Q1 & ", TipoDepHist = TipoDep"
   Q1 = Q1 & ", TipoDepLey21210Hist = TipoDepLey21210 "
   Q1 = Q1 & ", DepLey21256Hist = DepLey21256 "
   
   Q1 = Q1 & ", DepNormalHist = iif( TipoDep = " & DEP_NORMAL & ", DepNormalHist + DepNormal, 0)"
   Q1 = Q1 & ", DepAceleradaHist = iif( TipoDep = " & DEP_ACELERADA & ", DepAceleradaHist + DepAcelerada, 0)"
   Q1 = Q1 & ", DepInstantHist = iif( TipoDep = " & DEP_INSTANTANEA & ", DepInstantHist + DepInstant, 0)"
   Q1 = Q1 & ", DepDecimaParteHist = iif( TipoDep = " & DEP_DECIMAPARTE & ", DepDecimaParteHist + DepDecimaParte, 0)"
   Q1 = Q1 & ", DepDecimaParte2Hist = iif( TipoDep = " & DEP_DECIMAPARTE2 & ", DepDecimaParte2Hist + DepDecimaParte2, 0)"
   
   Q1 = Q1 & ", ValReajustadoNetoAnt = ValReajustadoNeto"
   Q1 = Q1 & ", DepAcumHist = DepAcumFinal"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   'limpiamos FExported en tabla nueva y actualizamos campos
   'actualizamos campos de depreciación en un query aparte para evitar que se pisen los cambios
   
   Q1 = "UPDATE MovActivoFijo SET "
   Q1 = Q1 & "  FExported = 0"
   Q1 = Q1 & ", FImported = " & CLng(Int(Now))
   Q1 = Q1 & ", TipoDep = iif(TipoDep = " & DEP_DECIMAPARTE & " or TipoDep = " & DEP_DECIMAPARTE2 & " or TipoDep = " & DEP_INSTANTANEA & ", " & DEP_NORMAL & ", TipoDep )"
   Q1 = Q1 & ", DepNormal = iif( TipoDep = " & DEP_NORMAL & " or TipoDep = " & DEP_DECIMAPARTE & " or TipoDep = " & DEP_DECIMAPARTE2 & " or TipoDep = " & DEP_INSTANTANEA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepAcelerada = iif( TipoDep = " & DEP_ACELERADA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepInstant = 0"
   Q1 = Q1 & ", DepDecimaParte = 0"
   Q1 = Q1 & ", DepDecimaParte2 = 0"
   Q1 = Q1 & ", TotalmenteDepreciado = iif(ValorLibro = 1, 1, TotalmenteDepreciado) "
   Q1 = Q1 & ", TipoDepLey21210 = 0 "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
            
   'Mensaje con cantidad
   MsgBox1 "Se importaron " & NImported & " Activos Fijos del año anterior, con vida útil residual o no depreciables.", vbInformation
   
   Call CloseRs(Rs)
   
   'actualizamos marca de exportación en tabla año anterior con fecha actual
   Q1 = "UPDATE MovActivoFijo "
   Q1 = Q1 & " SET FExported = " & CLng(Int(Now))
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Call ExecSQL(DbMain, Q1)
      
   GenActFijoResidualEmpJuntas = True

End Function

Public Function PrtPieBalance(PrtObj As Object, ByVal Pag As Integer, ByVal LeftX As Integer, ByVal RightX As Integer) As Integer
   Dim PrtPage As Object
   Dim CurY As Integer
   Dim BalFooter As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Col0 As Integer
   Dim Col1 As Integer
   Dim Col2 As Integer
   Dim Col3 As Integer
   Dim Col4 As Integer
   
   BalFooter = 2000
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   PrtPage.Print
   PrtPage.Print
   
   
   CurY = PrtPage.CurrentY
   
   If PrtPage.CurrentY >= PrtPage.Height - BalFooter - 1500 Then
      Call gPrtReportes.PrtFooter(PrtPage, "Continua >>>", RightX)
      Set PrtPage = NewPage(PrtObj)
      Pag = Pag + 1
      
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      CurY = PrtPage.CurrentY

   End If
   
   Q1 = "SELECT Contador, RutContador, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2 "
   Q1 = Q1 & " FROM Empresa WHERE Id =" & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      PrtPage.CurrentY = CurY + 500
      PrtPage.CurrentX = LeftX
      
      PrtPage.Print "Contador: ";
      If vFld(Rs("Contador")) <> "" Then
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("Contador"));
      End If
      
      PrtPage.CurrentX = LeftX + 5000
      If vFld(Rs("RutContador")) <> "" Then
         PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutContador")));
      Else
         PrtPage.Print "RUT: ";
      End If

      PrtPage.CurrentX = LeftX + 7000
      PrtPage.Print "Firma"
     
      
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
      
      
      
      PrtPage.Print "Rep. Legal: ";
      If vFld(Rs("RepLegal1")) <> "" Then
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("RepLegal1"));
      End If
      
      PrtPage.CurrentX = LeftX + 5000
      If vFld(Rs("RutRepLegal1")) <> "" Then
         PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutRepLegal1")));
      Else
         PrtPage.Print "RUT: ";
      End If

      PrtPage.CurrentX = LeftX + 7000
      PrtPage.Print "Firma"
         
      
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
      
      
      If vFld(Rs("RepLegal2")) <> "" Then
         PrtPage.Print "Rep. Legal: ";
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("RepLegal2"));
         
         PrtPage.CurrentX = LeftX + 5000
         If vFld(Rs("RutRepLegal2")) <> "" Then
            PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutRepLegal2")));
         Else
            PrtPage.Print "RUT: ";
         End If
   
         PrtPage.CurrentX = LeftX + 7000
         PrtPage.Print "Firma"
      End If
   
   End If
   
   Call CloseRs(Rs)
   
   If LCase(TypeName(PrtPage)) <> "picturebox" Then
      PrtPage.EndDoc
   End If
   
End Function

Public Function PrtPieBalanceFirma(PrtObj As Object, ByVal Pag As Integer, ByVal LeftX As Integer, ByVal RightX As Integer, ByVal TipoPrint As Integer) As Integer 'TipoPrint 0 = preview 1 = print
   Dim PrtPage As Object
   Dim CurY As Integer
   Dim BalFooter As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Col0 As Integer
   Dim Col1 As Integer
   Dim Col2 As Integer
   Dim Col3 As Integer
   Dim Col4 As Integer
   Dim vPath As String
   
   BalFooter = 2000
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   PrtPage.Print
   PrtPage.Print
   
   
   CurY = PrtPage.CurrentY
   
   If PrtPage.CurrentY >= PrtPage.Height - BalFooter - 1500 Then
      Call gPrtReportes.PrtFooter(PrtPage, "Continua >>>", RightX)
      Set PrtPage = NewPage(PrtObj)
      Pag = Pag + 1
      
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      PrtPage.Print
      CurY = PrtPage.CurrentY

   End If
   
   Q1 = "SELECT Contador, RutContador, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2 "
   Q1 = Q1 & " FROM Empresa WHERE Id =" & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      PrtPage.CurrentY = CurY + 500
      PrtPage.CurrentX = LeftX
      
      PrtPage.Print "Contador: ";
      If vFld(Rs("Contador")) <> "" Then
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("Contador"));
      End If
      
      PrtPage.CurrentX = LeftX + 5000
      If vFld(Rs("RutContador")) <> "" Then
         PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutContador")));
      Else
         PrtPage.Print "RUT: ";
      End If

      PrtPage.CurrentX = LeftX + 7000
      PrtPage.Print "Firma"
      'firma 2861570 tema 2
      
      vPath = ExisteFirma("Contador")

      If vPath <> "" Then

       If gEmpresa.RutContador <> "" Then
             PrtPage.CurrentX = LeftX + 7000
              PrtPage.PaintPicture LoadPicture(vPath), PrtPage.CurrentX + 500, PrtPage.CurrentY - 950

        End If
      End If
'      'fin 2861570
'
'       '2861570 tema 2
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
      'fin 2861570
      
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
       
      PrtPage.Print "Rep. Legal: ";
      If vFld(Rs("RepLegal1")) <> "" Then
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("RepLegal1"));
      End If
      
      PrtPage.CurrentX = LeftX + 5000
      If vFld(Rs("RutRepLegal1")) <> "" Then
         PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutRepLegal1")));
      Else
         PrtPage.Print "RUT: ";
      End If

      PrtPage.CurrentX = LeftX + 7000
      PrtPage.Print "Firma"
         
      '2861570 tema 2
      
       vPath = ExisteFirma("RepLegal1")

      If vPath <> "" Then
        If gEmpresa.RutRepLegal1 <> "" Then
           PrtPage.CurrentX = LeftX + 7000
           PrtPage.PaintPicture LoadPicture(vPath), PrtPage.CurrentX + 500, PrtPage.CurrentY - 1000
        End If

      End If

'      'fin 2861570
'
'       '2861570 tema 2
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
'      'fin 2861570

      
      PrtPage.CurrentY = PrtPage.CurrentY + 500
      PrtPage.CurrentX = LeftX
      
           
      If vFld(Rs("RepLegal2")) <> "" Then
         PrtPage.Print "Rep. Legal: ";
         PrtPage.CurrentX = LeftX + 1200
         PrtPage.Print "Sr(a/ta). " & vFld(Rs("RepLegal2"));
         
         PrtPage.CurrentX = LeftX + 5000
         If vFld(Rs("RutRepLegal2")) <> "" Then
            PrtPage.Print "RUT: " & FmtCID(vFld(Rs("RutRepLegal2")));
         Else
            PrtPage.Print "RUT: ";
         End If
   
         PrtPage.CurrentX = LeftX + 7000
         PrtPage.Print "Firma"
         
         '2861570 tema 2
         vPath = ExisteFirma("RepLegal2")

        If vPath <> "" Then
          If gEmpresa.RutRepLegal2 <> "" Then
           PrtPage.CurrentX = LeftX + 7000
           PrtPage.PaintPicture LoadPicture(vPath), PrtPage.CurrentX + 500, PrtPage.CurrentY - 770
          End If
        End If

      'fin 2861570
         
      End If
   
   End If
   
   Call CloseRs(Rs)
   
   If LCase(TypeName(PrtPage)) <> "picturebox" Then
      PrtPage.EndDoc
   End If
   
End Function

Public Sub CleanCtasBas(Ctas As CuentasBas_t)

   'Cuentas asociadas a IVA y Otros Impuestos
   Ctas.IdCtaIVACred = 0
   Ctas.IdCtaIVADeb = 0
   Ctas.IdCtaOtrosImpCred = 0
   Ctas.IdCtaOtrosImpDeb = 0
   
   'contracuenta pago facturas
   Ctas.IdCtaPagoFacturas = 0
   Ctas.IdCtaCobFacturas = 0
   
   'cuentas retenciones
   Ctas.IdCtaImpRet = 0
   Ctas.IdCtaNetoHon = 0
   Ctas.IdCtaNetoDieta = 0
   Ctas.IdCtaImpUnico = 0
   
   'cuentas de Patrimonio y Resultado Ejercicio
   Ctas.IdCtaPatrimonio = 0
   Ctas.IdCtaResEje = 0
   
   'cuenta de Crédito IVA para remanente año anterior
   Ctas.IdCtaCredIVA = 0
   
   'Cta de Retencion 3$ por préstamo
   Ctas.IdCtaRet3Porc = 0
   
   '2699582
   Ctas.IdCtaPpmObligatorio = 0
   Ctas.IdCtaPpmVoluntario = 0
   'FIN 2699582
   
   
End Sub

Public Function GetNombreTipoDoc(ByVal TipoLib As Integer, ByVal TipoDoc As Integer) As String
   Dim i As Integer
   
   GetNombreTipoDoc = ""
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).TipoDoc = TipoDoc Then
         GetNombreTipoDoc = gTipoDoc(i).Nombre
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetTipoDoc(ByVal TipoLib As Integer, ByVal TipoDoc As Integer) As Integer
   Dim i As Integer
   
   GetTipoDoc = -1
   
   If TipoLib = 0 Or TipoDoc = 0 Then
      Exit Function
   End If
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).TipoDoc = TipoDoc Then
         GetTipoDoc = i
         Exit Function
      End If
   
   Next i
   
End Function

Public Sub FillTipoDoc(Cb_TipoDoc As Control, ByVal TipoLib As Integer, ByVal CallClear As Boolean, ByVal AddBlankItem As Boolean)
   Dim i As Integer
   Dim InitTipoLib As Boolean
   
   If CallClear Then
      Cb_TipoDoc.Clear
   End If
   
   If AddBlankItem Then
   
      Cb_TipoDoc.AddItem " "
      Cb_TipoDoc.ItemData(Cb_TipoDoc.NewIndex) = 0
      
   End If
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib Then
         InitTipoLib = True
         Cb_TipoDoc.AddItem gTipoDoc(i).Nombre
         Cb_TipoDoc.ItemData(Cb_TipoDoc.NewIndex) = gTipoDoc(i).TipoDoc
         
      ElseIf InitTipoLib = True Then   'terminó el libro solicitado (están ordenados por TipoLib, TipoDoc)
         Exit For
            
      End If
   
   Next i
   
   If Cb_TipoDoc.ListCount > 0 Then
      Cb_TipoDoc.ListIndex = 0
   End If
   
   
End Sub

Public Function FindTipoDoc(ByVal TipoLib As Integer, ByVal Diminutivo As String) As Integer
   Dim i As Integer
   
   FindTipoDoc = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).Diminutivo = Diminutivo Then
         FindTipoDoc = gTipoDoc(i).TipoDoc
      End If
      
   Next i
   
End Function
'Retorna Indice de estructura gTipoValLib, correspondiente al TipoLib y TipoValLib que recibe como parámetro
Public Function FindTipoValLib(ByVal TipoLib As Integer, ByVal TipoValLib As Integer) As Integer
   Dim i As Integer
   
   FindTipoValLib = 0
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib And gTipoValLib(i).TipoValLib = TipoValLib Then
         FindTipoValLib = i
         Exit Function
      End If
      
   Next i
   
End Function
Public Function FindTipoDocLibCaja(ByVal Diminutivo As String, TipoLib As Integer) As Integer
   Dim i As Integer
   
   FindTipoDocLibCaja = 0
   
   'en el libro de caja se juntan los egresos que corresponden a Libro de Compras y Libro de Retenciones
   'Dado que los diminutivos son únicos, no se usa TipoLib. Más aún se retorna el valor de TipoLib
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).Diminutivo = Diminutivo Then
         FindTipoDocLibCaja = gTipoDoc(i).TipoDoc
         TipoLib = gTipoDoc(i).TipoLib
      End If
      
   Next i
   
End Function
Public Sub FillTipoValLib(Cb_TipoValLib As Control, ByVal TipoLib As Integer, ByVal CallClear As Boolean, ByVal AddBlankItem As Boolean, Optional ByVal Atributo As String = "", Optional ByVal TipoDoc As Integer = 0, Optional ByVal OcultarImpAdicDescontinuados As Boolean = False)
   Dim i As Integer
   Dim InitTipoLib As Boolean
   
   If CallClear Then
      Cb_TipoValLib.Clear
   End If
   
   If AddBlankItem Then
   
      Cb_TipoValLib.AddItem " "
      Cb_TipoValLib.ItemData(Cb_TipoValLib.NewIndex) = 0
      
   End If
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib Then
      
         InitTipoLib = True
         
         If TipoDoc = 0 Or (TipoDoc <> 0 And (gTipoValLib(i).TipoDoc = "" Or InStr(gTipoValLib(i).TipoDoc, "," & TipoDoc & ",") <> 0)) Then
            If Atributo = "" Or (Atributo <> "" And gTipoValLib(i).Atributo = Atributo) Then
               If Not OcultarImpAdicDescontinuados Or (OcultarImpAdicDescontinuados And Not gTipoValLib(i).Descontinuado) Then
                  Cb_TipoValLib.AddItem gTipoValLib(i).Nombre
                  Cb_TipoValLib.ItemData(Cb_TipoValLib.NewIndex) = gTipoValLib(i).TipoValLib
               End If
            End If
         End If
      ElseIf InitTipoLib = True Then   'terminó el libro solicitado (están ordenados por TipoLib, TipoValLib (Codigo))
         Exit For
            
      End If
   
   Next i
   
   If Cb_TipoValLib.ListCount > 0 Then
      Cb_TipoValLib.ListIndex = 0
   End If
     
End Sub
Public Sub FillClsTipoValLib(Cb As ClsCombo, ByVal TipoLib As Integer, ByVal CallClear As Boolean, ByVal AddBlankItem As Boolean, Optional ByVal Atributo As String = "", Optional ByVal TipoDoc As Integer = 0, Optional ByVal OcultarImpAdicDescontinuados As Boolean = False)
   Dim i As Integer
   Dim InitTipoLib As Boolean
   Dim Tasa As Single, EsRecuperable As Boolean
   Dim IdCuenta As Long
   
   If CallClear Then
      Call Cb.Clear
   End If
   
   If AddBlankItem Then
   
      Call Cb.AddItem(" ", 0, 0, 0)
      
   End If
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib Then
      
         InitTipoLib = True
         
         If TipoDoc = 0 Or (TipoDoc <> 0 And (gTipoValLib(i).TipoDoc = "" Or InStr(gTipoValLib(i).TipoDoc, "," & TipoDoc & ",") <> 0)) Then
            If Atributo = "" Or (Atributo <> "" And gTipoValLib(i).Atributo = Atributo) Then
               If Not OcultarImpAdicDescontinuados Or (OcultarImpAdicDescontinuados And Not gTipoValLib(i).Descontinuado) Then
                  IdCuenta = GetCtaImpAdic(TipoLib, gTipoValLib(i).TipoValLib, Tasa, EsRecuperable)
                  Call Cb.AddItem(gTipoValLib(i).Nombre, gTipoValLib(i).TipoValLib, gTipoValLib(i).CodSIIDTE, IdCuenta)
               End If
            End If
         End If
      ElseIf InitTipoLib = True Then   'terminó el libro solicitado (están ordenados por TipoLib, TipoValLib (Codigo))
         Exit For
            
      End If
   
   Next i
   
   If Cb.ListCount > 0 Then
      Cb.ListIndex = 0
   End If
     
End Sub
Public Function GetNombreTipoValLib(ByVal TipoLib As Integer, ByVal TipoValLib As Integer) As String
   Dim i As Integer
   
   GetNombreTipoValLib = ""
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib And gTipoValLib(i).TipoValLib = TipoValLib Then
         GetNombreTipoValLib = gTipoValLib(i).Nombre
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetTipoDocFromCodDocDTESII(ByVal TipoLib As Integer, ByVal CodDocDTESII As String) As Integer
   Dim i As Integer
   
   GetTipoDocFromCodDocDTESII = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And Val(gTipoDoc(i).CodDocDTESII) = Val(CodDocDTESII) Then
         GetTipoDocFromCodDocDTESII = gTipoDoc(i).TipoDoc
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetTipoDocFromCodDocSII(ByVal TipoLib As Integer, ByVal CodDocSII As String) As Integer
   Dim i As Integer
   
   GetTipoDocFromCodDocSII = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And Val(gTipoDoc(i).CodDocSII) = Val(CodDocSII) Then
         GetTipoDocFromCodDocSII = gTipoDoc(i).TipoDoc
         Exit Function
      End If
   
   Next i
   
End Function

Public Function GetTipoValLib(ByVal TipoLib As Integer, ByVal TipoValLib As Integer) As Integer
   Dim i As Integer
   
   GetTipoValLib = -1
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib And gTipoValLib(i).TipoValLib = TipoValLib Then
         GetTipoValLib = i
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetTipoTipoValLibFromCodSIIDTE(ByVal TipoLib As Integer, ByVal CodSIIDTE As String) As Integer
   Dim i As Integer
   
   GetTipoTipoValLibFromCodSIIDTE = -1
   
   For i = 0 To UBound(gTipoValLib)
   
      If gTipoValLib(i).TipoLib = TipoLib And gTipoValLib(i).CodSIIDTE = CodSIIDTE Then
         GetTipoTipoValLibFromCodSIIDTE = i
         Exit Function
      End If
   
   Next i
   
End Function
Public Function GetCtaImpAdic(ByVal TipoLib As Integer, ByVal TipoValLib As Integer, Tasa As Single, EsRecuperable As Boolean) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   If TipoLib = 0 Or TipoValLib = 0 Then
      Exit Function
   End If
   
   Q1 = "SELECT IdCuenta, Tasa, EsRecuperable FROM ImpAdic WHERE TipoLib = " & TipoLib & " AND TipoValor = " & TipoValLib
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   GetCtaImpAdic = 0
   Tasa = 0
   EsRecuperable = 0

   If Not Rs.EOF Then
      GetCtaImpAdic = vFld(Rs("IdCuenta"))
      Tasa = vFld(Rs("Tasa"))
      EsRecuperable = vFld(Rs("EsRecuperable"))
   End If
   
   Call CloseRs(Rs)
   
End Function
'PS
Public Function UpdateUltUsado(PapelFoliado As Boolean, nFolio As Integer) As Boolean
   Dim Q1 As String
   
   UpdateUltUsado = False
   If ChkOpt(gEmpresa.Opciones, OPT_ACTUSADO) And PapelFoliado Then
      If gFoliacion.UltUsado + nFolio > gFoliacion.UltTimbrado Then
         MsgBox1 "No se pudo actualizar último folio usado, porque el último folio timbrado es menor ", vbExclamation
         Exit Function
      End If
      
      Q1 = "UPDATE Timbraje SET "
      Q1 = Q1 & " UltUsado =" & gFoliacion.UltUsado + nFolio
      Q1 = Q1 & ", FUltUsado=" & SqlNowI(DbMain)
      Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
      gFoliacion.UltUsado = gFoliacion.UltUsado + nFolio
      gFoliacion.FUltUsado = Int(Now)
      
      UpdateUltUsado = True
   End If
      
End Function
Public Function DeleteComprobante(ByVal idcomp As Long, Optional ByVal MsgPregunta As Boolean = True, Optional ByVal DocFull As Boolean = False) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Fecha As Long
   Dim Tipo As Integer
   Dim CorrComp As Long
   Dim tblDoc As String
   Dim tblCom As String
   Dim tblMovCom As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String

   DeleteComprobante = False
   
'   If DocFull Then
'    tblDoc = "DocumentoFull"
'    tblCom = "ComprobanteFull"
'    tblMovCom = "MovComprobanteFull"
'   Else
    tblDoc = "Documento"
    tblCom = "Comprobante"
    tblMovCom = "MovComprobante"
'   End If
   If idcomp = 0 Then
      Exit Function
   End If
      
   If Not ValidaIngresoComp(False, True) Then
      Exit Function
   End If
   
   Q1 = "SELECT Fecha, Tipo, Correlativo FROM " & tblCom & " WHERE IdComp=" & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Fecha = vFld(Rs("Fecha"))
      Tipo = vFld(Rs("Tipo"))
      CorrComp = vFld(Rs("Correlativo"))
   End If
   
   Call CloseRs(Rs)
   
   If Tipo = TC_APERTURA Then
   
      MsgBox1 "No es posible eliminar el comprobante de apertura. Puede anularlo, si lo desea.", vbExclamation + vbOKOnly
      Exit Function
      
   End If
      
   If GetEstadoMes(month(Fecha)) <> EM_ABIERTO Then
      MsgBox1 "No es posible eliminar este comprobante. La fecha de emisión no corresponde a un mes actualmente abierto.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   'preguntamos si está seguro
   If MsgPregunta Then
      If MsgBox1("¿Está seguro que desea eliminar el comprobante de " & gTipoComp(Tipo) & " N° " & CorrComp & "?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   'eliminar el comprobante => debemos soltar los documentos asociados al comprobante
   
   'desconciliamos los movimientos (si hubiera)
'   Q1 = "UPDATE DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
'   Q1 = Q1 & " AND DetCartola.IdEmpresa = MovComprobante.IdEmpresa AND DetCartola.Ano = MovComprobante.Ano"
'   Q1 = Q1 & " SET DetCartola.IdMov = 0 WHERE MovComprobante.IdComp = " & IdComp
'   Q1 = Q1 & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " DetCartola "
   sFrom = " DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
   sFrom = sFrom & " AND DetCartola.IdEmpresa = MovComprobante.IdEmpresa AND DetCartola.Ano = MovComprobante.Ano"
   sSet = " DetCartola.IdMov = 0 "
   sWhere = " WHERE MovComprobante.IdComp = " & idcomp
   sWhere = sWhere & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
   'feña
'   sFrom = Replace(sFrom, "MovComprobante", tblMovCom)
'   sWhere = Replace(sWhere, "MovComprobante", tblMovCom)
   ' fin feña
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

   'soltamos cuotas de pago si corresponde
   Q1 = "UPDATE DocCuotas SET Estado = " & ED_PENDIENTE & ", IdCompPago = 0"
   Q1 = Q1 & " WHERE IdCompPago = " & idcomp & " AND Estado = " & ED_PAGADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
         
   'los docs. que vienen de centralización los dejamos pendientes
   Q1 = "UPDATE " & tblDoc & " SET IdCompCent = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE IdCompCent = " & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'los docs. que vienen de pago autom.: dejamos en estado ED_CENTRALIZADO si tiene IdCompCent <> 0
   Q1 = "UPDATE " & tblDoc & " SET IdCompPago = 0, Estado = " & ED_CENTRALIZADO & ", SaldoDoc = NULL WHERE IdCompPago = " & idcomp & " AND IdCompCent <> 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'los docs. que vienen de pago autom.: dejamos pendientes si tiene IdCompCent = 0 (esto no debiera ocurrir nunca, pero por si las moscas)
   Q1 = "UPDATE " & tblDoc & " SET IdCompPago = 0, Estado = " & ED_PENDIENTE & ", SaldoDoc = NULL WHERE IdCompPago = " & idcomp & " AND IdCompCent = 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
         
   If Not DocFull Then
    Q1 = "UPDATE MovDocumento SET IdCompCent = 0, IdCompPago = 0 WHERE IdCompCent = " & idcomp & " OR IdCompPago = " & idcomp
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    Call ExecSQL(DbMain, Q1)
   End If
   'eliminamos el comprobante
'   Q1 = "DELETE * FROM MovComprobante WHERE IdComp = " & IdComp
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE IdComp = " & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   'Call DeleteSQL(DbMain, "MovComprobante", Q1)
   'FEÑA
    Call DeleteSQL(DbMain, tblMovCom, Q1)
   'FIN FEÑA
'   Q1 = "DELETE * FROM Comprobante WHERE IdComp = " & IdComp
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE IdComp = " & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   'Call DeleteSQL(DbMain, "Comprobante", Q1)
   'feña
   Call DeleteSQL(DbMain, tblCom, Q1)
   'fin feña
   DeleteComprobante = True

End Function
Public Sub GetResIVA(ByVal Mes As Integer, ByVal Ano As Integer, TotIVACred As Double, TotIVADeb As Double, TotIEPDGen As Double, TotIEPDTransp As Double)
   Dim Q1 As String
   Dim Rs As Recordset
   
   TotIVACred = 0
   TotIVADeb = 0
   
   Q1 = "SELECT Documento.TipoLib, EsRebaja, Sum(IVA) As SumIVA "
   Q1 = Q1 & " FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Documento.TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & ")"
   Q1 = Q1 & " AND " & SqlYearLng("FEmision") & " = " & Ano

   If Mes > 0 Then
      Q1 = Q1 & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Documento.TipoLib, EsRebaja"
   Q1 = Q1 & " ORDER BY Documento.TipoLib, EsRebaja"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      If vFld(Rs("TipoLib")) = LIB_COMPRAS Then
         If vFld(Rs("EsRebaja")) <> 0 Then
            TotIVACred = TotIVACred - vFld(Rs("SumIVA"))
         Else
            TotIVACred = TotIVACred + vFld(Rs("SumIVA"))
         End If
      
      Else
         If vFld(Rs("EsRebaja")) <> 0 Then
            TotIVADeb = TotIVADeb - vFld(Rs("SumIVA"))
         Else
            TotIVADeb = TotIVADeb + vFld(Rs("SumIVA"))
         End If
      
      End If
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   
   TotIEPDGen = 0
   TotIEPDTransp = 0
   
   Q1 = "SELECT Documento.TipoLib, EsRebaja, IdTipoValLib, Sum(Debe-Haber) As SumMov "
   Q1 = Q1 & " FROM (Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc) "
   Q1 = Q1 & " INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "Documento")
   Q1 = Q1 & " WHERE Documento.TipoLib = " & LIB_COMPRAS
'   Q1 = Q1 & " AND MovDocumento.IdTipoValLib IN(" &       & "," & LIBCOMPRAS_IMPESPPETRCARGACF & ")"    'valores eliminados
   Q1 = Q1 & " AND MovDocumento.IdTipoValLib IN(" & LIBCOMPRAS_IMPESPDIESEL & "," & LIBCOMPRAS_IMPESPDIESELTRANS & ")"
   Q1 = Q1 & " AND MovDocumento.EsRecuperable <> 0 "
   Q1 = Q1 & " AND " & SqlYearLng("FEmision") & " = " & Ano

   If Mes > 0 Then
      Q1 = Q1 & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Documento.TipoLib, EsRebaja, IdTipoValLib"
   Q1 = Q1 & " ORDER BY Documento.TipoLib, EsRebaja"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      If vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IMPESPDIESEL Then
         If vFld(Rs("EsRebaja")) <> 0 Then
            TotIEPDGen = TotIEPDGen - Abs(vFld(Rs("SumMov")))
         Else
            TotIEPDGen = TotIEPDGen + Abs(vFld(Rs("SumMov")))
         End If
      
      ElseIf vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IMPESPDIESELTRANS Then
         If vFld(Rs("EsRebaja")) <> 0 Then
            TotIEPDTransp = TotIEPDTransp - Abs(vFld(Rs("SumMov")))
         Else
            TotIEPDTransp = TotIEPDTransp + Abs(vFld(Rs("SumMov")))
         End If
      
      End If
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
      
End Sub
'Public Function GetRemIVAUTM(ByVal Mes As Integer, ByVal Ano As Integer, RemUTMMes As Double) As Integer
'   Dim TotIVACred As Double
'   Dim TotIVADeb As Double
'   Dim Remanente As Double
'   Dim ValUTM As Double
'   Dim Fecha As Long
'   Dim RemUTM As Double
'   Dim RemUTMAnoAnt As Double
'   Dim TotRemUTM As Double
'   Dim i As Integer, Rc As Integer
'   Dim IVAIrrec As Double
'   Dim TotIEPDGen As Double, TotIEPDTransp As Double
'
'   RemUTMMes = 0
'
'   GetRemIVAUTM = 0
'   Rc = GetRemIVAAnoAnt(RemUTMAnoAnt)
'   If Rc < 0 Then
'      GetRemIVAUTM = Rc
'      Exit Function
'   End If
'
'   'OJO redondeamos a dos decimales
'   RemUTMAnoAnt = Format(RemUTMAnoAnt, DBLFMT2)
'
'   TotRemUTM = RemUTMAnoAnt
'
'   For i = 1 To Mes - 1
'
'      Fecha = DateSerial(Ano, i + 1, 1)            'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011
'
'      Call GetResIVA(i, Ano, TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
'      IVAIrrec = GetIVAIrrec(i, Ano)
'
'      Remanente = TotIVACred + TotIEPDGen + TotIEPDTransp - IVAIrrec - TotIVADeb
'
'      If GetValMoneda("UTM", ValUTM, Fecha, True) = True And ValUTM > 0 Then
'         RemUTM = Remanente / ValUTM
'
'         'OJO redondeamos a dos decimales
'         RemUTM = vFmt(Format(RemUTM, DBLFMT2))
'
'         TotRemUTM = TotRemUTM + RemUTM
'
'         If TotRemUTM < 0 Then
'            TotRemUTM = 0
'         End If
'
''      ElseIf Remanente > 0 Then
'      Else
'         MsgBox1 "No se encontró el valor de la UTM para el mes de " & gNomMes(Month(Fecha)) & " " & Ano, vbExclamation
'         GetRemIVAUTM = ERR_VALUTM
'         Exit For
'
'      End If
'
'
'   Next i
'
'   RemUTMMes = TotRemUTM
'
'End Function


Public Function GetRemIVAUTM(ByVal Mes As Integer, ByVal Ano As Integer, RemUTMMes As Double) As Integer
   Dim TotIVACred As Double
   Dim TotIVADeb As Double
   Dim Remanente As Double
   Dim ValUTM As Double
   Dim Fecha As Long
   Dim RemUTM As Double
   Dim RemUTMAnoAnt As Double
   Dim TotRemUTM As Double
   Dim i As Integer, Rc As Integer
   Dim IVAIrrec As Double
   Dim TotIEPDGen As Double, TotIEPDTransp As Double
   Dim IVARetParcial As Double, IVARetTotal As Double
   Dim AjusteIVAMensual As Double
   
   RemUTMMes = 0
   
   GetRemIVAUTM = 0
   Rc = GetRemIVAAnoAnt(RemUTMAnoAnt)
   If Rc < 0 Then
      GetRemIVAUTM = Rc
      Exit Function
   End If
      
   'OJO redondeamos a dos decimales
   RemUTMAnoAnt = Format(RemUTMAnoAnt, DBLFMT2)
   
   TotRemUTM = RemUTMAnoAnt
   
   For i = 1 To Mes - 1
  
      Fecha = DateSerial(Ano, i + 1, 1)            'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011
      
      Call GetResIVA(i, Ano, TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
      IVAIrrec = GetIVAIrrec(i, Ano, i = 1) ' 15 feb 2020
      Call GetIVARet(i, Ano, IVARetParcial, IVARetTotal, False) ' 15 feb 2010
      
      AjusteIVAMensual = GetAjusteIVAMensual(i)    ' FCA 28/09/2021
           
      Remanente = TotIVACred + TotIEPDGen + TotIEPDTransp + AjusteIVAMensual - IVAIrrec - (TotIVADeb - IVARetParcial - IVARetTotal)
            
      If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
         RemUTM = (Remanente / ValUTM)
         
         'OJO redondeamos a dos decimales
         
         '3321695
         RemUTM = Format(RemUTM, DBLFMT2)
          'RemUTM = (Fix(RemUTM * 100)) / 100
         '3321695
         TotRemUTM = TotRemUTM + RemUTM
         
         If TotRemUTM < 0 Then
            TotRemUTM = 0
         End If
         
      Else
         MsgBox1 "No se encontró el valor de la UTM para el mes de " & gNomMes(month(Fecha)) & " " & Ano, vbExclamation
         GetRemIVAUTM = ERR_VALUTM
         Exit For
      End If
         
      
   Next i
   
   RemUTMMes = TotRemUTM
   
End Function
Public Function GetAjusteIVAMensual(ByVal Mes As Integer) As Double    ' FCA 28/09/2021
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetAjusteIVAMensual = 0
   
   If Mes <= 0 Then
      Exit Function
   End If
   
   Q1 = "SELECT Valor FROM AjusteIVAMensual WHERE Mes = " & Mes & " AND Ano = " & gEmpresa.Ano & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetAjusteIVAMensual = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
      
End Function


Public Function GetIVAIrrec(ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal bCreaTmp As Boolean = 1) As Double
   Dim Where As String
   Dim ResOImp() As ResOImp_t
   Dim i As Integer
   
   GetIVAIrrec = 0
   
   Where = " " & SqlYearLng("FEmision") & " = " & Ano
   
   If Mes > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   
'   If lTipoLib > 0 Then
'      Where = Where & " AND Documento.TipoLib = " & lTipoLib
'   End If
   
   Call GenResOImp(Where, ResOImp, , bCreaTmp)
      
   For i = 0 To UBound(ResOImp)
      
      If ResOImp(i).TipoLib = LIB_COMPRAS And ResOImp(i).CodValLib <> 0 Then
                                    
         If ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC1 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC2 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC3 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC4 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC9 Then
            GetIVAIrrec = ResOImp(i).Valor
         End If
         
      Else
         Exit For
      
      End If
   Next i
   
   
End Function

Public Function GetIVARet(ByVal Mes As Integer, ByVal Ano As Integer, IVARetParcial As Double, IVARetTotal As Double, Optional ByVal bCreaTmp As Boolean = 1) As Integer
   Dim Where As String
   Dim ResOImp() As ResOImp_t
   Dim i As Integer
   
   GetIVARet = False
   
   IVARetParcial = 0
   IVARetTotal = 0
   
   Where = " " & SqlYearLng("FEmision") & " = " & Ano
   
   If Mes > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   
'   If lTipoLib > 0 Then
'      Where = Where & " AND Documento.TipoLib = " & lTipoLib
'   End If
   
   Call GenResOImp(Where, ResOImp, , bCreaTmp) ' 15 feb 2020
      
   For i = 0 To UBound(ResOImp)
      
      If ResOImp(i).CodValLib <> 0 Then
                                    
         If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_PARCIAL Then
            IVARetParcial = ResOImp(i).Valor
         End If
         
         If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_TOTAL Then
            IVARetTotal = ResOImp(i).Valor
         End If
         
      Else
         Exit For
      
      End If
   Next i
   
   GetIVARet = True
   
End Function


Public Function GetRemIVAAnoAnt(RemUTMAnoAnt As Double) As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   'Dim idCuentaCredIVA As Long
   Dim Saldo As Double
   Dim Fecha As Long
   Dim ValUTM As Double
   
   'idCuentaCredIVA = 0
   RemUTMAnoAnt = 0
   GetRemIVAAnoAnt = 0
   
   'se obtiene de valor remanente de año anterior, almacenado en el cierre del año o en la apertura de este año por ingreso directo
   Q1 = "SELECT RemIVAUTMAnoAnt FROM EmpresasAno WHERE idEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      RemUTMAnoAnt = vFld(Rs("RemIVAUTMAnoAnt"))
       '3426794
      RemIVAAnoAnt = True
   Else
      RemIVAAnoAnt = False
      '3426794
   End If

   Call CloseRs(Rs)
   
   
   
   
   
      
   'se elimina a solicitud de Cristofer 17 junio 2015
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CTACREDIVA'"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Not Rs.EOF Then
'      idCuentaCredIVA = vFld(Rs("Valor"))
'   End If
'
'   Call CloseRs(Rs)
'
'   If idCuentaCredIVA = 0 Then
'      MsgBox1 "No se ha definido la cuenta de Crédito de IVA" & vbCrLf & "para obtener el remanente del año anterior.", vbExclamation
'      GetRemIVAAnoAnt = ERR_DEFCUENTA
'      Exit Function
'   End If
'
'   Q1 = "SELECT  Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
'   Q1 = Q1 & " FROM (MovComprobante "
'   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta) "
'   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp"
'   Q1 = Q1 & " WHERE Comprobante.Tipo = " & TC_APERTURA & " AND MovComprobante.IdCuenta =" & idCuentaCredIVA
'   Q1 = Q1 & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
'   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   Saldo = 0
'
'   If Not Rs.EOF Then
'      Saldo = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
'   End If
'
'   Call CloseRs(Rs)
'
'   If Saldo > 0 Then
'
'      'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011)
'      'Fecha = DateSerial(gEmpresa.Ano - 1, 12, 1)
'      Fecha = DateSerial(gEmpresa.Ano, 1, 1)
'
'      If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then 'UTM diciemnbre año anterior
'         RemUTMAnoAnt = Saldo / ValUTM
'
'      Else
'         MsgBox1 "No se encontró el valor de la UTM para el mes de " & gNomMes(Month(Fecha)) & " " & gEmpresa.Ano - 1, vbExclamation
'         GetRemIVAAnoAnt = ERR_VALUTM
'         Exit Function
'      End If
'
'   End If
   
   
End Function
Public Sub UpdParamEmpresa(ByVal Tipo As String, ByVal Codigo As Integer, ByVal Valor As String)
   Dim Rs As Recordset, Q1 As String, oValor As String, Rc As Long
   
   Tipo = UCase(Tipo)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='" & Tipo & "' AND Codigo=" & Codigo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      oValor = vFld(Rs("Valor"))
      Call CloseRs(Rs)
      If oValor <> Valor Then
         Q1 = "UPDATE ParamEmpresa SET Valor='" & ParaSQL(Left(Valor, 30)) & "' WHERE Tipo='" & Tipo & "' AND Codigo=" & Codigo
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Rc = ExecSQL(DbMain, Q1)
      End If
   Else
      Call CloseRs(Rs)
      
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('" & Tipo & "'," & Codigo & ",'" & ParaSQL(Left(Valor, 30)) & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
      Rc = ExecSQL(DbMain, Q1)
   
   End If

End Sub

Public Function GetParamEmpresa(ByVal Tipo As String, Optional ByVal Codigo As Integer = 0) As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetParamEmpresa = ""

   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='" & Tipo & "'"
   If Codigo <> 0 Then
      Q1 = Q1 & " AND Codigo = " & Codigo
   End If
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      GetParamEmpresa = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
End Function

'Si Ano = 0 se asume que es año actual y por lo tanto no se leen los indices, se usan los que están
'Si no, se leen los indices del año Ano y al final se releen los de gEmpresa.Ano
Public Function RecalcDepResidual(Optional ByVal Ano As Integer = 0, Optional ByVal Msg As Boolean = True) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim ValInit As Double
   Dim TieneCredArt33 As Integer
   Dim YaDep As Integer
   Dim ADep As Integer
   Dim Residual As Integer
   Dim FMes As Integer
   Dim FAño As Integer
   Dim Factor As Double
   Dim ValReajustado As Double
   Dim CredArt33 As Double
   Dim ValReajCred As Double
   Dim DepAcumAct As Double
   Dim ValDepreciar As Double
   Dim DepMensual As Double
   Dim DepPeriodo As Double
   Dim DepAcumuladaAno As Double
   Dim ValLibro As Double
   Dim MsgFact As Boolean
   Dim DispResid As Integer
   Dim IdxAno As Integer
   
   RecalcDepResidual = True
   
   If Ano <> 0 Then
      IdxAno = Ano
      Call ReadIndices(Ano)  'leemos indices Ano
   Else
      IdxAno = gEmpresa.Ano
   End If
   
   Q1 = "SELECT IdActFijo, MovActivoFijo.IdDoc, MovActivoFijo.IdComp, IdMovComp, TipoMovAF, MovActivoFijo.Fecha, MovActivoFijo.FechaVentaBaja "
   Q1 = Q1 & ", MovActivoFijo.Cantidad, MovActivoFijo.Descrip, MovActivoFijo.Neto, MovActivoFijo.IVA, Cred4Porc"
   Q1 = Q1 & ", VidaUtil, DepNormal, DepAcelerada, DepInstant, DepDecimaParte, TipoDep, DepNormalHist, DepAceleradaHist, DepInstantHist, DepDecimaParteHist, TipoDepHist "
   Q1 = Q1 & ", DepAcumHist, MovActivoFijo.IdCuenta "
   Q1 = Q1 & "  FROM MovActivoFijo "
   Q1 = Q1 & "  WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & "  ORDER BY MovActivoFijo.Fecha, MovActivoFijo.Descrip"

   Set Rs = OpenRs(DbMain, Q1)
      
   Do While Rs.EOF = False
   
      If vFld(Rs("FechaVentaBaja")) = 0 Or vFld(Rs("FechaVentaBaja")) > DateSerial(IdxAno, 12, 31) Then
      
         ValInit = vFld(Rs("Neto"))
         
         TieneCredArt33 = IIf(vFld(Rs("Cred4Porc")) <> 0, 1, 0)
                     
         Select Case vFld(Rs("TipoDepHist"))
            Case DEP_NORMAL
               YaDep = vFld(Rs("DepNormalHist"))
            Case DEP_ACELERADA
               YaDep = vFld(Rs("DepAceleradaHist"))
            Case DEP_INSTANTANEA
               YaDep = vFld(Rs("DepInstantHist"))
            Case DEP_DECIMAPARTE
               YaDep = vFld(Rs("DepDecimaParteHist"))
            Case Else
               YaDep = 0
         End Select
         
         DispResid = vFld(Rs("VidaUtil")) - YaDep
         
         Select Case vFld(Rs("TipoDep"))
            Case DEP_NORMAL
               ADep = vFld(Rs("DepNormal"))
            Case DEP_ACELERADA
               ADep = vFld(Rs("DepAcelerada"))
            Case DEP_INSTANTANEA
               ADep = vFld(Rs("DepInstant"))
            Case DEP_DECIMAPARTE
               ADep = vFld(Rs("DepDecimaParte"))
         End Select
         
         Residual = DispResid - ADep   'actualizar en DB
         If Residual < 0 Then
            MsgBeep (vbExclamation)
         End If
         
         Q1 = "UPDATE MovActivoFijo SET "
'         Q1 = Q1 & "  DepAcumFinal = " & DepAcumuladaAno    'este valor se graba en el reporte de activo fijo para luego se lleva al año siguiente en la función GenActFijoResidual
         Q1 = Q1 & " VidaUtilResidual = " & Residual
         Q1 = Q1 & " WHERE IdActFijo = " & vFld(Rs("IdActFijo"))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call ExecSQL(DbMain, Q1)
      
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If Ano <> 0 Then
      Call ReadIndices     'leemos indices año gEmpresa.Ano
   End If

End Function
'PS 9/2/2006 Por el error q se produjo con un cliente q el prog. guardaba mal el id del comprobante de apertura
'agregué este código
Public Sub CheckCompAPertura()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdCompAper As Long
   
   Q1 = "SELECT IdCompAper, NCompAper "
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      IdCompAper = vFld(Rs("IdCompAper"))
   End If
   Call CloseRs(Rs)
      
   Q1 = "SELECT idComp FROM Comprobante WHERE Tipo=" & TC_APERTURA
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If IdCompAper <> vFld(Rs("idComp")) Then
         IdCompAper = vFld(Rs("idComp"))
         
         Q1 = "UPDATE EmpresasAno "
         Q1 = Q1 & " SET idCompAper=" & IdCompAper
         Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.id
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
   Call CloseRs(Rs)

End Sub

Private Sub VerificaMultiCompApertura()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim NAper As Long
   Dim DelFrom As String, DelWhere As String

   'primero verificamos múltiples comprobantes de apertura financieros
   Q1 = "SELECT Count(*) FROM Comprobante WHERE Tipo =" & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   NAper = 0
   If Rs.EOF = False Then
      NAper = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   If NAper > 0 Then    'ya existe uno o más comprobantes de apertura, algo raro pasó
                        'los eliminamos todos para que luego se vuelva a generar
      DelFrom = " MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      DelFrom = DelFrom & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      DelWhere = " WHERE Comprobante.Tipo =" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_FINANCIERO
      DelWhere = DelWhere & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano

      Call DeleteJSQL(DbMain, "MovComprobante", DelFrom, DelWhere)
                       
      Q1 = " WHERE Tipo =" & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "Comprobante", Q1)
   
   End If
   
   'luego verificamos múltiples comprobantes de apertura tributarios
   Q1 = "SELECT Count(*) FROM Comprobante WHERE Tipo =" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   NAper = 0
   If Rs.EOF = False Then
      NAper = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   If NAper > 0 Then    'ya existe uno o más comprobantes de apertura, algo raro pasó
                        'los eliminamos todos para que luego se vuelva a generar
                       
      DelFrom = " MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      DelFrom = DelFrom & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      DelWhere = " WHERE Comprobante.Tipo =" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
      DelWhere = DelWhere & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano

      Call DeleteJSQL(DbMain, "MovComprobante", DelFrom, DelWhere)
         
      Q1 = " WHERE Tipo =" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "Comprobante", Q1)
   
   
   End If

End Sub
Public Function HayAnoAnterior() As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim InitAno As String
   
   'veamos si la empresa tiene historia (año anterior)
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='INITAÑO'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      InitAno = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   If InitAno = "EMPHISTORIA" Then
      HayAnoAnterior = True
   Else
      HayAnoAnterior = False
   End If

End Function


Public Function QryNContrib(ByVal Qry As String, EntidadHR() As EntidadHR_t) As Long
#If DATACON = 1 Then       'Access

'   Dim Conn As ADODB.Connection
'   Dim Rs As ADODB.Recordset
   Dim HrConnStr As String
   Dim DbPath As String
   Dim TblName1 As String, TblName2 As String, TblName3 As String
   Dim Rs As Recordset
   Dim Rc As Long
   Dim i As Integer

   Rc = 0
   QryNContrib = 0

   If Not gLinkF22 Then
      Exit Function
   End If

   On Error Resume Next

   DbPath = gHRPath & "\PAR\BD_HR_admin.mdb"
   HrConnStr = ";PWD=200803hr;"

   TblName1 = "Adm_NContrib"
   Rc = LinkMdbTable(DbMain, DbPath, TblName1, "HR_" & TblName1, , False, HrConnStr)
   TblName2 = "Adm_Comuna"
   Rc = LinkMdbTable(DbMain, DbPath, TblName2, "HR_" & TblName2, , False, HrConnStr)
   TblName3 = "Adm_Region"
   Rc = LinkMdbTable(DbMain, DbPath, TblName3, "HR_" & TblName3, , False, HrConnStr)


'   Set Conn = New ADODB.Connection
'   DbPath = gHRPath & "\PAR\BD_HR_admin.mdb"
'   HrConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbPath & ";Jet OLEDB:Database Password=200803hr;"
'
'   Call Conn.Open(HrConnStr)

   If ERR Then
      'MsgBox "Error H" & Hex(Err) & ", " & Error & NL & DbPath, vbExclamation
      QryNContrib = -1
      Exit Function
   End If

'   Set Rs = Conn.Execute(Qry)

   Set Rs = OpenRs(DbMain, Qry)

'   If Rs Is Nothing Then
'      Conn.Close
'      Set Conn = Nothing
'      Exit Function
'   End If

   ReDim EntidadHR(10)
   i = 0

   Do While Not Rs.EOF
      EntidadHR(i).Nombre = Trim(vFld(Rs("NC_Nombre")) & " " & vFld(Rs("NC_Paterno")) & " " & vFld(Rs("NC_Materno")))

      EntidadHR(i).Direccion = vFld(Rs("NC_Calle"), True) & " #" & vFld(Rs("NC_Nro"))

      If Trim(vFld(Rs("NC_Depto"))) <> "" Then
         EntidadHR(i).Direccion = EntidadHR(i).Direccion & " dpto. " & vFld(Rs("NC_Depto"), True)
      End If

      EntidadHR(i).NombreCorto = vFld(Rs("NC_NomCorto"), True)

      EntidadHR(i).Ciudad = vFld(Rs("NC_Ciudad"), True)

      EntidadHR(0).Region = vFld(Rs("Reg_Orden"))
      EntidadHR(0).Comuna = vFld(Rs("Com_Nombre"))

      EntidadHR(i).Tel = vFld(Rs("NC_Fono"))
      EntidadHR(i).Fax = vFld(Rs("NC_Fax"))
      EntidadHR(i).email = vFld((Rs("NC_Correo")))
      EntidadHR(i).DirPostal = vFld(Rs("NC_Dir_Postal"))

      i = i + 1

      If UBound(EntidadHR) <= i Then
         ReDim Preserve EntidadHR(i + 10)
      End If

      Rs.MoveNext
   Loop

   If i > 0 Then
      ReDim Preserve EntidadHR(i - 1)
      Rc = i
   End If

   QryNContrib = Rc

'   Rs.Close
'   Set Rs = Nothing
'   Conn.Close
'   Set Conn = Nothing

   Call CloseRs(Rs)

   Call ExecSQL(DbMain, "DROP TABLE " & TblName1)
   Call ExecSQL(DbMain, "DROP TABLE " & TblName2)
   Call ExecSQL(DbMain, "DROP TABLE " & TblName3)
   
#End If

End Function


Public Sub DesConciliarMov(ByVal idMov As Long)
   Dim Q1 As String

   'desconciliamos los movimientos (si hubiera)
   Q1 = "UPDATE DetCartola SET idMov=0 WHERE IdMov = " & idMov
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE MovComprobante SET idCartola=0 WHERE IdMov = " & idMov
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)

End Sub

Public Function GetNumDocVSD(ByVal TipoLib As Integer, ByVal TipoDoc As Integer) As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Max(NumDoc) FROM Documento"
   Q1 = Q1 & " WHERE TipoLib = " & TipoLib & " AND TipoDoc = " & TipoDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
      GetNumDocVSD = 1
   Else                              ' se supone que siempre es numérico porque es asignado por el sistema mismo
      GetNumDocVSD = Val(vFld(Rs(0))) + 1
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function GetCodSIIDoc(ByVal IdDoc As Long, NumDoc As String) As String   'Retorna CodDocSII o CodDocDTESII
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT DTE, CodDocSII, CodDocDTESII, NumDoc FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc"
   Q1 = Q1 & " WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If vFld(Rs("DTE")) <> 0 Then
         GetCodSIIDoc = vFld(Rs("CodDocDTESII"))
      Else
         GetCodSIIDoc = vFld(Rs("CodDocSII"))
      End If
      NumDoc = vFld(Rs("NumDoc"))
   Else
      GetCodSIIDoc = ""
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function GetNumDocVSDLibCaja(ByVal TipoLib As Integer, ByVal TipoDoc As Integer) As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Max(NumDoc) FROM LibroCaja "
   Q1 = Q1 & " WHERE TipoLib = " & TipoLib & " AND TipoDoc = " & TipoDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
      GetNumDocVSDLibCaja = 1
   Else                              ' se supone que siempre es numérico porque es asignado por el sistema mismo
      GetNumDocVSDLibCaja = Val(vFld(Rs(0))) + 1
   End If
   
   Call CloseRs(Rs)
   
End Function


Public Sub AddLogComprobantes(ByVal idcomp As Long, ByVal IdUsuario As Long, ByVal IdOper As Integer, ByVal Fecha As Double, ByVal Estado As Integer, Optional ByVal DCorrelativoComp As Long = 0, Optional ByVal DFechaComp As Long = 0, Optional ByVal DTipoComp As Integer = 0, Optional ByVal DEstadoComp As Integer = 0, Optional ByVal DTipoAjuste As Integer = 0)
   Dim Q1 As String
      
   Q1 = "INSERT INTO LogComprobantes( IdComp, IdUsuario, IdOper, Fecha, Estado, CorrComp, FechaComp, TipoComp, EstadoComp, TipoAjusteComp, IdEmpresa, Ano ) "
   Q1 = Q1 & " VALUES( " & idcomp & "," & IdUsuario & "," & IdOper & "," & str(CDbl(Fecha)) & "," & Estado & "," & DCorrelativoComp & "," & DFechaComp & "," & DTipoComp & "," & DEstadoComp & "," & DTipoAjuste & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
   
   Call ExecSQL(DbMain, Q1)

End Sub
         
Public Function GetEntRelacionada(ByVal IdEntidad As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   If IdEntidad = 0 Then
      GetEntRelacionada = 0
   End If
   
   Q1 = "SELECT EntRelacionada FROM Entidades "
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   GetEntRelacionada = False
   If Not Rs.EOF Then
      GetEntRelacionada = vFld(Rs("EntRelacionada"))
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function GetFactorCM(ByVal AnoMes As Long) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   
   GetFactorCM = 0
   
   Q1 = "SELECT fCM FROM IPC WHERE AnoMes = " & AnoMes
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      GetFactorCM = vFld(Rs("fCM"))
      
   End If
   
   Call CloseRs(Rs)
      

End Function

Public Sub MsgLey21210(Optional ByVal MsgSp As String = "")
   Dim Frm As FrmMsgConBreak
   Dim Msg As String
   
   Set Frm = New FrmMsgConBreak
      
   If MsgSp <> "" Then
      Msg = MsgSp
   Else
      Msg = gMsgLey21210
   End If
   
   Call Frm.FView(Msg, "NoDispMsgLey21210")
      
   Set Frm = Nothing

End Sub


'pipe 2850275

Public Sub EmpresasLpRemu()
Dim PathDbLpRemu As String
Dim ConnStr As String
Dim q1Remu As String

    Dim PathDbLpContab As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   Dim Q1 As String
   Dim RsDao As dao.Recordset

    Dim FNBaseAccess2 As String
   Dim DbAccess2 As Database
   Dim Q2 As String
   Dim RsDao2 As dao.Recordset
   Dim RsDao3 As dao.Recordset
   Dim ConnStr2 As String
   Dim bErrMsg As String

   Dim Rs1 As Recordset
   Dim Rs2 As Recordset

    Dim Rc As Long
    
    Dim codComuna As Integer
    
    codComuna = 0

'2850275

   If gDbType = SQL_ACCESS Then



  If OpenlDbRemu() = True Then

   'leemos la lista de empresas
   q1Remu = "SELECT Rut,Direccion,email,Telefono,Fax,ciudades.CodCiudad,Web,ciudades.CodRegion,empresas.CodComuna,RutRep,NomRep,Giro,ciudades.Ciudad,ltrim(rtrim(Comunas.Comuna)) as nomComuna FROM Empresas,Ciudades,Comunas  "
     q1Remu = q1Remu & " where ciudades.codciudad = empresas.codciudad and empresas.CodComuna = Comunas.CodComuna and Ciudades.CodCiudad = Comunas.CodCiudad and rut ='" & gEmpresa.Rut & "'"

   Set RsDao = OpenRsDao(lDbRemu, q1Remu)

   Do While RsDao.EOF = False

    'veamos si existe archivo LPContab.mdb en el path de la aplicación
   PathDbLpContab = gDbPath & "\LPContab.mdb"
   If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPContab.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpContab & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then

         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPContab.mdb.", vbExclamation
         Exit Sub

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta

         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess2) = vbCancel Then

            Exit Sub

         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpContab = FNBaseAccess2

            If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPContab.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Sub

            End If

         End If

      End If

   End If

    '2868088
   ConnStr2 = ";PWD=" & PASSW_LEXCONT & ";"
   'ConnStr2 = ";PWD=" & PASSW_LEXCONT_NEW & ";"
   '2868088
   Set DbAccess2 = OpenDatabase(PathDbLpContab, False, False, ConnStr2)

   If ERR <> 0 Or DbAccess2 Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpContab.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Sub
   End If

   'gFrmMain.MousePointer = vbHourglass

   'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
   Set RsDao2 = OpenRsDao(DbAccess2, Q1)
   
        Q2 = ""
        Q2 = " select id,comuna from regiones "
        Q2 = Q2 & " where comuna = Ucase('" & TextoSinAcentos(vFldDao(RsDao("nomComuna"))) & "')"
        
        Set RsDao3 = OpenRsDao(DbAccess2, Q2)
          
        
        If Not RsDao3.EOF Then
        codComuna = vFld(RsDao3("id"))
        End If
        

   If Not RsDao2.EOF Then
       Q1 = "UPDATE Empresa SET "
            Q1 = Q1 & " Calle='" & vFldDao(RsDao("Direccion")) & "'"
            Q1 = Q1 & ", EMail='" & vFldDao(RsDao("email")) & "'"
            Q1 = Q1 & ", Telefonos='" & vFldDao(RsDao("Telefono")) & "'"
            Q1 = Q1 & ", Fax='" & vFldDao(RsDao("Fax")) & "'"
            Q1 = Q1 & ", Ciudad='" & vFldDao(RsDao("Ciudad")) & "'"
            Q1 = Q1 & ", Web='" & vFldDao(RsDao("Web")) & "'"
            Q1 = Q1 & ", Region=" & vFldDao(RsDao("CodRegion"))
            Q1 = Q1 & ", Comuna=" & codComuna
            Q1 = Q1 & ", ComunaPostal=" & codComuna
            Q1 = Q1 & ", RutRepLegal1='" & vFldDao(RsDao("RutRep")) & "'"
            Q1 = Q1 & ", RepLegal1='" & vFldDao(RsDao("NomRep")) & "'"
            Q1 = Q1 & ", Giro='" & vFldDao(RsDao("Giro")) & "'"
            Q1 = Q1 & " WHERE Id =" & gEmpresa.id & " AND Ano =" & gEmpresa.Ano

            Call ExecSQL(DbMain, Q1, False)

    End If

    RsDao.MoveNext
   Call CloseRs(RsDao2)
   Call CloseRs(RsDao3)
   
   Call CloseDb(DbAccess2)
   Loop

   Call CloseRs(RsDao)

   Call CloseDb(DbAccess)


    End If

   ElseIf gDbType = SQL_SERVER Then

    If OpenMsSqlRemu() = True Then

   q1Remu = "SELECT Rut,Direccion,email,Telefono,Fax,ciudades.CodCiudad,Web,ciudades.CodRegion,empresas.CodComuna,RutRep,NomRep,Giro,ciudades.Ciudad,ltrim(rtrim(Comunas.Comuna)) as nomComuna FROM Empresas,Ciudades,Comunas  "
     q1Remu = q1Remu & " where ciudades.codciudad = empresas.codciudad and empresas.CodComuna = Comunas.CodComuna and Ciudades.CodCiudad = Comunas.CodCiudad and rut ='" & gEmpresa.Rut & "'"

   Set RsDao = OpenRsDao(lDbRemu, q1Remu)

   Do While RsDao.EOF = False
     'leemos la lista de empresas
        Q1 = ""
        Q1 = "SELECT * FROM Empresa WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
        Set Rs1 = OpenRs(DbMain, Q1)
            
            
        Q2 = ""
        Q2 = " select id,comuna from regiones "
        Q2 = Q2 & " where comuna = UPPER('" & TextoSinAcentos(vFldDao(RsDao("nomComuna"))) & "')"
        
        Set Rs2 = OpenRs(DbMain, Q2)
        
       
        
        
        If Not Rs2.EOF Then
        codComuna = vFld(Rs2("id"))
        End If
        
        Call CloseRs(Rs2)
            
        If Not Rs1.EOF Then
            Q1 = "UPDATE Empresa SET "
            Q1 = Q1 & " Calle='" & vFldDao(RsDao("Direccion")) & "'"
            Q1 = Q1 & ", EMail='" & vFldDao(RsDao("email")) & "'"
            Q1 = Q1 & ", Telefonos='" & vFldDao(RsDao("Telefono")) & "'"
            Q1 = Q1 & ", Fax='" & vFldDao(RsDao("Fax")) & "'"
            Q1 = Q1 & ", Ciudad='" & vFldDao(RsDao("Ciudad")) & "'"
            Q1 = Q1 & ", Web='" & vFldDao(RsDao("Web")) & "'"
            Q1 = Q1 & ", Region=" & vFldDao(RsDao("CodRegion"))
            Q1 = Q1 & ", Comuna=" & codComuna
            Q1 = Q1 & ", ComunaPostal=" & codComuna
            Q1 = Q1 & ", RutRepLegal1='" & vFldDao(RsDao("RutRep")) & "'"
            Q1 = Q1 & ", RepLegal1='" & vFldDao(RsDao("NomRep")) & "'"
            Q1 = Q1 & ", Giro='" & vFldDao(RsDao("Giro")) & "'"
            Q1 = Q1 & " WHERE Id =" & gEmpresa.id & " AND Ano =" & gEmpresa.Ano

            Call ExecSQL(DbMain, Q1, False)

         End If

         RsDao.MoveNext
        Call CloseRs(Rs1)
        

   Loop
 Call CloseRs(RsDao)

    End If

   End If

   'fin 2850275

End Sub

'2850275

'para Access
Private Function OpenlDbRemu() As Boolean
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim Cfg As String
   Dim AuxPathlDbRemu As String
   Dim lIdEmpresaRem As Long
   Dim Idx As Integer
   Dim Qry As QueryDef
   Dim RazonSoc As String
   Dim Buf As String
   Dim TmpQry As String
   Dim lEmpSep As Boolean


   OpenlDbRemu = False

   lPathlDbRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   Buf = GetIniString(gIniFile, "Config", "VersionRemu", "")
   lRemuSQLServer = False

   If lPathlDbRemu = "" Then
      'MsgBox1 "Falta configurar la localización del sistema de Remuneraciones. Utilice la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If

   If Not ExistFile(lPathlDbRemu) Then
      MsgBox1 "No se encontró el archivo de " & IIf(lRemuSQLServer, "configuración", "base dee datos") & " del sistema de Remuneraciones." & vbCrLf & vbCrLf & "Verifique la localización del archivo en la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If

   If lRemuSQLServer Then
      If OpenMsSqlRemu() = False Then
         Exit Function
      End If

   Else

#If DATACON = 1 Then
      Cfg = GetIniString(gCfgFile, "Config", "Secur", "")
      If Cfg <> SG_SEGCFG Then         'suponenemos que si la base de LPContab no tiene clave, tampoco la tiene la de FairPay
         lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
      Else
         lConnStr = ""
      End If

      lConnStr = Mid(lConnStr, 2)

#Else
      lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"    'si estamos en SQLServer, se asume que la base de Remu Access tiene  clave

#End If

      On Error Resume Next

      Set lDbRemu = OpenDatabase(lPathlDbRemu, False, False, lConnStr)

      If lDbRemu Is Nothing Then
         Call AddLog("OpenDbRemu: Error " & ERR & ", " & Error & ", " & lPathlDbRemu)

         lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"    'asumimos que faltó la clave (caso particular de Fairware)

         Set lDbRemu = OpenDatabase(lPathlDbRemu, False, False, lConnStr)


         If lDbRemu Is Nothing Then
            'MsgBox1 "No se pudo abrir la base de datos " & vbCrLf & lPathlDbRemu & vbCrLf & vbCrLf & Error, vbExclamation
            Call AddLog("OpenDbRemu: Error " & ERR & ", " & Error & ", " & lPathlDbRemu)
            Exit Function
         End If
      End If

   End If

   'en remuneraciones se permite tener dos empresas con el mismo RUT, por eso se da la opción al cliente de seleccionar cuál de las dos (9 abril 2019)
'   Q1 = "SELECT IdEmpresa, RazonSoc FROM Empresas WHERE Rut = '" & gEmpresa.Rut & "'"
'   'Q1 = "SELECT IdEmpresa, RazonSoc FROM Empresas "
'   Set Rs = OpenRsDao(lDbRemu, Q1)
'   If Not Rs.EOF Then
'      lIdEmpresaRem = vFldDao(Rs("IdEmpresa"))
'      RazonSoc = vFldDao(Rs("RazonSoc"))
'
'      Rs.MoveNext
'
'      If Not Rs.EOF Then
'         If MsgBox1("Existen dos empresas con este mismo RUT." & vbCrLf & "¿Desea obtener los datos de '" & RazonSoc & "'?", vbQuestion + vbYesNo) = vbNo Then
'            lIdEmpresaRem = vFldDao(Rs("IdEmpresa"))
'         End If
'      End If
'
'   End If
'   Call CloseRs(Rs)

'   If lIdEmpresaRem = 0 Then
'      MsgBox "No se ha encontrado esta empresa en el sistema de Remuneraciones.", vbExclamation + vbOKOnly
'      Exit Function
'   End If

'   If Not lRemuSQLServer Then
'
'      Q1 = "SELECT Codigo FROM Param WHERE Tipo = 'EMPSEP'"
'      Set Rs = OpenRsDao(lDbRemu, Q1)
'      If Rs.EOF = False Then
'         lEmpSep = IIf(vFldDao(Rs("Codigo")) <> 0, True, False)
'      Else
'         lEmpSep = False
'      End If
'      Call CloseRs(Rs)
'
'      lDbRemu.Close
'      Set lDbRemu = Nothing
'
'      AuxPathlDbRemu = lPathlDbRemu
'
'
'      If lEmpSep Then
'         Idx = InStrRev(AuxPathlDbRemu, "\")
'         If Idx > 0 Then
'            AuxPathlDbRemu = Left(AuxPathlDbRemu, Idx)
'         End If
'         AuxPathlDbRemu = AuxPathlDbRemu & "Empresas\" & gEmpresa.Rut & "_" & lIdEmpresaRem & ".mdb"
'      End If
'
'' se descomenta para quitar la clave a la base de datos.
''        Set lDbRemu = OpenDatabase(AuxPathlDbRemu, True, False, lConnStr)
''        lDbRemu.NewPassword SG_PASSW_FAIRPAY, ""
''        Call CloseDb(lDbRemu)
''        lConnStr = ""
''        Set lDbRemu = OpenDatabase(AuxPathlDbRemu, False, False, lConnStr)
''      If lDbRemu Is Nothing Then
''         Call AddLog("OpenDbRemu : Error " & ERR & ", " & Error & ", " & AuxPathlDbRemu)
''         Exit Function
''      End If
'
'   End If

   OpenlDbRemu = True

End Function
'fin 2850275

'2850275
' Para MsSql Server
' Verificar SHOW VARIABLES LIKE 'lower_case_table_names'  que sea 1 o 2
Function OpenMsSqlRemu() As Boolean
   Dim Rc As Integer, SqlPort As Long, Usr As String, Psw As String, i As Integer
   Dim ConnStr As String, Host As String, UsrPsw As String, DbName As String
   Dim sErr1 As Long, sError1 As String, Encript As Boolean, CfgFile As String

   On Error Resume Next

   OpenMsSqlRemu = False
   lPathlDbRemu = GetIniString(gIniFile, "Config", "PathRemu", "")

   If Not lDbRemu Is Nothing Then
      lDbRemu.Close
      Set lDbRemu = Nothing
   End If

   CfgFile = lPathlDbRemu
   If LCase(Right(lPathlDbRemu, 10)) = "lpremu.cfg" Then
      lEsLPRemu = True
   ElseIf LCase(Right(lPathlDbRemu, 11)) = "fairpay.cfg" Then
      lEsLPRemu = False
   Else
     ' MsgBox1 "Falta especificar correctamente el archivo de configuración de Remuneraciones." & vbCrLf & "Utilice la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If

   Host = Trim(GetIniString(CfgFile, "MS Sql", "Host", ""))

   If Host = "" Then
      MsgBox1 "Falta especificar el servidor de base de datos." & vbCrLf & "Comuníquese con su administrador.", vbCritical
      Exit Function
   End If

   SqlPort = Val(GetIniString(CfgFile, "MS Sql", "Port", "1433"))

   If lEsLPRemu Then
      Debug.Print "Db lpremu=" & FwEncrypt1("               lpremu             ", 56516)
      DbName = GetIniString(CfgFile, "MS Sql", "DB", FwDecrypt1("6E2C6B2B6C2E71357A40874F98E2D8D7DFDBDA5E2F8154287D532A825B35906C4927", 56516))

      Usr = GetIniString(CfgFile, "MS Sql", "User", "lp" & "re" & "mu")
   Else
      Debug.Print "Db fairpay=" & FwEncrypt1("           fairpay           ", 56516)
      DbName = GetIniString(CfgFile, "MS Sql", "DB", FwDecrypt1("9053975C2269317A448F5BA89DABABAAB3C553287E552D86603B977452", 56516))

      Usr = GetIniString(CfgFile, "MS Sql", "User", "fai" & "rp" & "ay")
   End If


   Debug.Print "Hola Psw=" & FwEncrypt1("     " & DbName & "   #" & "      hola       ", 731982) ' ojo con el #
   Debug.Print "Oficial Psw=" & FwEncrypt1("     " & DbName & "   #" & "     _F&].[r94%.        ", 731982) ' ojo con el #

   Psw = GetIniString(CfgFile, "MS Sql", "Pswk")

   If Psw = "" Then
      MsgBox1 "Falta especificar la clave del servidor de base de datos de Remuneraciones." & vbCrLf & "Comuníquese con su administrador.", vbCritical
      Exit Function
   End If

   Psw = Trim(FwDecrypt1(Psw, 731982))
   i = InStr(Psw, "#")
   Psw = Trim(Mid(Psw, i + 1))

   UsrPsw = "U" & "ID=" & Usr & ";P" & "WD=" & Psw & ";"

   ConnStr = "Driver={SQL Server};Server=" & Host & ";MARS_Connection=yes;Database=" & DbName & ";" ' 2 abr 2018

   On Error Resume Next

   Set lDbRemu = OpenDatabase("", False, False, ConnStr & UsrPsw)

'   Set lDbRemu = New Connection
'   lDbRemu.ConnectionString = ConnStr & UsrPsw
'   lDbRemu.Open

   If ERR Then
      If ERR <> 3059 Then
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & ConnStr, vbCritical
      End If
      Call AddLog("OpenMsSqlRemu: Error " & ERR & ", " & Error & ", " & ConnStr)

      Set lDbRemu = Nothing

      End
      Exit Function
   End If

   If ERR Then
      sErr1 = ERR.Number
      sError1 = ERR.Description
      MsgErr "Verifique que esté bien definido el servidor de la base de datos y que tenga los privilegios necesarios."
      Call AddLog("Error " & sErr1 & ", " & sError1 & ", [" & ConnStr & "]")
   Else
      OpenMsSqlRemu = True

      If Psw = "" Then
         Psw = GetConnectInfo(lDbRemu, "PWD")
         UsrPsw = "User=" & Usr & ";PWD=" & Psw & ";"
      End If

'      gConnStr = ConnStr & UsrPsw   ' Para la exportación

'      lDbRemuDate = GetDbNow(lDbRemu)

   End If

End Function

'2850275
Public Function TextoSinAcentos(ByVal Texto As String) As String
' Esta función devuelve el texto sin acentos
Dim lngTexto As Long
Dim i As Long
Dim strCaracter As String * 1
Dim strNormalizado As String

lngTexto = Len(Texto)
If lngTexto = 0 Then
TextoSinAcentos = ""
Exit Function
End If
For i = 1 To lngTexto
strCaracter = Mid(Texto, i, 1)
Select Case strCaracter
Case "Á", "À", "Â", "Ä", "Ã"
strCaracter = "A"
Case "á", "à", "â", "ä", "ã"
strCaracter = "a"
Case "É", "È", "Ê", "Ë"
strCaracter = "E"
Case "é", "è", "ê", "ë"
strCaracter = "e"
Case "Í", "Ì", "Î", "Ï"
strCaracter = "I"
Case "í", "ì", "î", "ï"
strCaracter = "i"
Case "Ó", "Ò", "Ô", "Ö", "Õ"
strCaracter = "O"
Case "ó", "ò", "ô", "ö", "õ"
strCaracter = "o"
Case "Ú", "Ù", "Û", "Ü"
strCaracter = "U"
Case "ú", "ù", "û", "ü"
strCaracter = "u"
Case "Ý"
strCaracter = "Y"
Case "ý", "ÿ"
strCaracter = "y"
End Select
TextoSinAcentos = TextoSinAcentos & strCaracter
Next i
End Function

'fin  2850275

'2861570
Public Function ExisteFirma(ByVal Tipo As String) As String
Dim Q1 As String
Dim Rs As Recordset

ExisteFirma = ""

Q1 = ""
Q1 = "Select * from Firmas where idEmpresa =" & gEmpresa.id
Q1 = Q1 & " And Tipo ='" & Tipo & "'"
Set Rs = OpenRs(DbMain, Q1)

If Rs.EOF = False Then
ExisteFirma = vFld(Rs("patch"))

End If
Call CloseRs(Rs)
End Function
'2861570



#If DATACON = 1 Then
Public Function RecalcSaldosAnoAnterior(ByVal vDbMain As Database, ByVal IdEmpresa As Long, ByVal Ano As Integer, Optional ByVal bIniNull As Boolean = 1)
#Else
Public Function RecalcSaldosAnoAnterior(ByVal vDbMain As ADODB.Connection, ByVal IdEmpresa As Long, ByVal Ano As Integer, Optional ByVal bIniNull As Boolean = 1)
#End If

   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   'PIPE otros doc
   Dim Rs2 As Recordset
   Dim Debe As Double
   Dim Haber As Double
   Dim Saldo As Double
   Dim CurIdDoc As Long
   Dim WhLib As String
   
   '2931541
   Dim PagoAnoAnterior As Double
   '2931541
   
   '3025162
   Dim RsCopy As Object
   '3025162
   
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim DocOtroEsCargo As Boolean
   Dim SetPagado As Boolean
   
   If IdEmpresa = 0 Then
      IdEmpresa = gEmpresa.id
   End If
   
   If Ano = 0 Then
      Ano = gEmpresa.Ano
   End If
   
   
   WhLib = " Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") "
   
   'marcamos notas de crédito y débito con SaldoDoc = NULL que tienen factura asociada con SaldoDoc = NULL
'   Q1 = "UPDATE Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
'   Q1 = Q1 & " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
'   Q1 = Q1 & " SET Documento.SaldoDoc = NULL WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
'   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   Tbl = " Documento "
   sFrom = " Documento INNER JOIN Documento as Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
   sFrom = sFrom & JoinEmpAno(gDbType, "Documento", "Documento_1", True)  ' " AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano"
   sSet = " Documento.SaldoDoc = NULL "
   sWhere = " WHERE " & WhLib & " AND Documento_1.SaldoDoc IS NULL"
   sWhere = sWhere & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano
'   Call ExecSQL(DbMain, Q1)
   If bIniNull Then ' 14 feb 2020: ya se le puso NULL justo antes de llamar
      Call UpdateSQL(vDbMain, Tbl, sSet, sFrom, sWhere)
   End If
               
               
   'consulta de docs que no están enlazados a ningún comprobante
   Q1 = " SELECT 1, Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
   Q1 = Q1 & " Sum(MovDocumento.Debe) As Debe, Sum(MovDocumento.Haber) As Haber "
   
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541
   
   Q1 = Q1 & " FROM ((Documento  "
   Q1 = Q1 & "  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento", True) & " )" ' 14 feb 2020: se agrega , True
   'Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc"
   Q1 = Q1 & "  LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
'   Q1 = Q1 & "   AND Documento.IdEmpresa = vMovCompIdDoc.IdEmpresa AND Documento.Ano = vMovCompIdDoc.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc", True) & " )" ' 14 feb 2020
   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
'   Q1 = Q1 & "   AND Documento.IdEmpresa = Documento_1.IdEmpresa AND Documento.Ano = Documento_1.Ano "
   
     '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541

   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1", True) ' 14 feb 2020
  
   Q1 = Q1 & " WHERE " & WhLib & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL) "
   'Q1 = Q1 & "  AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND Documento.Estado <> " & ED_ANULADO
         'tomamos los que no están enlazados a un comprobante y los que están marcados como centralizados pero no tienen comprobante asociado (docs pendientes del año anterior)
   'Q1 = Q1 & "  AND (MovComprobante.IdComp IS NULL "
   Q1 = Q1 & "  AND (vMovCompIdDoc.IdDoc IS NULL "
   Q1 = Q1 & "  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   
   Q1 = Q1 & " GROUP BY Documento.TipoLib, Documento.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
   
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   
   Q1 = Q1 & " UNION "

   'consulta de movs. comprobantes que tienen docs enlazados
   Q1 = Q1 & " SELECT 2, Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, 0 as MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo, "
   Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) As Haber "

     '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541


   Q1 = Q1 & " FROM ((MovComprobante INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & "  INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & "  LEFT JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc  "
   '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1")

   Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO
   'Q1 = Q1 & " AND Documento.IdEntidad > 0 AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & "  AND " & WhLib
   Q1 = Q1 & "  AND Documento.SaldoDoc IS NULL "
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento.TipoLib, MovComprobante.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, Documento.DocOtroEsCargo "
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541

   Q1 = Q1 & " UNION "

   'consulta de docs ASOCIADOS que no están enlazados a ningún comprobante
   Q1 = Q1 & " SELECT 3, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0, "
   Q1 = Q1 & " Sum(MovDocumento.Debe) AS Debe, Sum(MovDocumento.Haber) AS Haber"
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541

   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
   '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "vMovCompIdDoc")


   Q1 = Q1 & " WHERE " & WhLib & " AND Documento.IdDocAsoc <> 0"
   Q1 = Q1 & " AND (MovDocumento.EsTotalDoc <> 0  OR MovDocumento.IdDoc IS NULL)"
   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
   Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND (vMovCompIdDoc.IdDoc IS NULL  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, MovDocumento.IdMovDoc, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
    '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541
   Q1 = Q1 & " UNION "

   'consulta de movs. comprobantes que tienen docs ASOCIADOS enlazados

   Q1 = Q1 & " SELECT 4, Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotPagadoAnoAnt, 0 AS MovIdDoc, Documento.IdDocAsoc, Documento_1.Estado As EstadoDocAsoc, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision, 0,  "
   Q1 = Q1 & " Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa "
   '2931541

   Q1 = Q1 & " FROM ((Documento INNER JOIN Documento AS Documento_1 ON Documento.IdDocAsoc = Documento_1.IdDoc "
    '2931541
   Q1 = Q1 & " And Documento.IdEmpresa = Documento_1.IdEmpresa"
   '2931541
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Documento_1") & " )"
   Q1 = Q1 & " INNER JOIN MovComprobante ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")

   Q1 = Q1 & " WHERE Comprobante.Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND " & WhLib
   Q1 = Q1 & " AND Documento.SaldoDoc IS NULL"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & IdEmpresa & " AND Documento.Ano = " & Ano

   Q1 = Q1 & " GROUP BY Documento_1.TipoLib, Documento_1.IdDoc, Documento.Total, Documento.Estado, Documento.TotpagadoAnoAnt, Documento.IdDocAsoc, Documento_1.Estado, Documento.IdCompCent, Documento.IdCompPago, Documento.FEmision "
   '2931541
   Q1 = Q1 & " ,Documento.IdEmpresa"
   '2931541

   Q1 = Q1 & " ORDER BY IdDoc"

   Set Rs = OpenRs(vDbMain, Q1)
      
  Do While Rs.EOF = False
        
      'detalle doc
      If CurIdDoc <> vFld(Rs("IdDoc")) Then   ' puede venir más de una vez cuando hay documentos asociados
      
        ' If vFld(Rs("IdDoc")) = 483 Then   'OJO CON EL ESTADO DEL DOCUEMENTO DEL AñO ANTERIOR QUE DEBE ESTAR CENTRALIZADO O PAGADO
        If vFld(Rs("IdDoc")) = 1297 Then

'299 --  300
            MsgBeep vbExclamation
         End If

         Debe = 0
         Haber = 0
         SetPagado = False
         
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
         
            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
               Saldo = vFld(Rs("Total"))
            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = (Debe - Haber)
                        
            End If
            
            '3047309
'            If vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
'               If Abs(vFld(Rs("Total"))) = Abs(Saldo) Then
'                 Saldo = Saldo + Abs(vFld(Rs("TotPagadoAnoAnt"))) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'               Else
'
'                Saldo = Abs(vFld(Rs("Total")) - (Saldo + vFld(Rs("TotPagadoAnoAnt")))) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
'               End If
'        Else
            
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt")) 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            'Saldo = Saldo - (IIf(vFld(Rs("TotPagadoAnoAnt")) > 0, vFld(Rs("TotPagadoAnoAnt")) * -1, vFld(Rs("TotPagadoAnoAnt"))))
        'End If
           '  3047309

'            Saldo = Saldo = Saldo - Abs(vFld(Rs("TotPagadoAnoAnt")))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            
            
            '2931541
            PagoAnoAnterior = vFld(Rs("TotPagadoAnoAnt"))
            '2931541
         Else
         
         
'            If (IsNull(Rs("Debe")) And IsNull(Rs("Haber"))) Or IsNull(Rs("MovIdDoc")) Then     'viene sin MovDocumento
'               Saldo = vFld(Rs("Total"))
'            Else
               Debe = vFld(Rs("Debe"))
               Haber = vFld(Rs("Haber"))
               Saldo = Debe - Haber
            
               If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                  If Saldo <= 0 Then
                     DocOtroEsCargo = True
                  Else
                     DocOtroEsCargo = False
                  End If
               End If
            
'            End If
           
           
            If IsNull(Rs("TotPagadoAnoAnt")) = False Then    'FCA 04 feb 2020: se asume que sólo el primer año el TotPagadoAnoAnt es NULL, por lo tanto, al segundo año se le agrega el total
               Saldo = Saldo + IIf(vFld(Rs("DocOtroEsCargo")), -1, 1) * vFld(Rs("Total"))
            End If
            
           
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0 "
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
           
            Call ExecSQL(vDbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
                              
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado, pero si ponemos el DocOtroEsCargo
                           
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               Else   'Saldo = Total
               
                  If vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano Then   'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
                                 
               End If
               
            Else
               Q1 = Q1 & ", DocOtroEsCargo = " & Abs(DocOtroEsCargo)
                                      
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
           
            Call ExecSQL(vDbMain, Q1)
         
         End If
         
      Else
                 
         If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then
            
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
           '3047309
            If vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
                Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt")) '- PagoAnoAnterior 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            Else
               Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt")) - PagoAnoAnterior 'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
            End If
            '3047309
            
            '2931541
         Else
            Debe = Debe + vFld(Rs("Debe"))
            Haber = Haber + vFld(Rs("Haber"))
            Saldo = Debe - Haber
            
            If vFld(Rs("Total")) = Debe Or vFld(Rs("Total")) = Haber Then      'para asegurarnos que se le asigna cargo o abono el primer uso, que debería ser el total del otro doc
                If Saldo <= 0 Then
                   DocOtroEsCargo = True
                Else
                   DocOtroEsCargo = False
                End If
             End If
                                   
            Saldo = Saldo - vFld(Rs("TotPagadoAnoAnt"))   'se resta valor pagado año anterior, si este doc proviniera de año anterior y tuviera pago parcial
         
         End If
         
         
         
         If vFld(Rs("IdDocAsoc")) <> 0 And vFld(Rs("IdDocAsoc")) <> vFld(Rs("IdDoc")) Then      'es una nota de crédito o débito asociada a una factura, ponemos saldo = 0 ya que este documento ya se considera en el cálculo del saldo de la factura
            Q1 = "UPDATE Documento SET SaldoDoc = 0"
            
            If vFld(Rs("EstadoDocAsoc")) = ED_PAGADO And vFld(Rs("Estado")) = ED_CENTRALIZADO Then   'la factura asociada está pagada y la nota de crédito o débito, está centralizada => marcamos como pagada
              Q1 = Q1 & ", Estado = " & ED_PAGADO
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            
            Call ExecSQL(vDbMain, Q1)

         Else
            Q1 = "UPDATE Documento SET SaldoDoc = " & Saldo
            
            'por si el usuario hizo un pago manual, sin la opción de generar pago automático
            If vFld(Rs("TipoLib")) <> LIB_OTROS And vFld(Rs("TipoLib")) <> LIB_REMU Then    'para LIB_OTROS no cambiamos estado
               
               If Abs(Saldo) <> vFld(Rs("Total")) Then
                  If vFld(Rs("Estado")) = ED_CENTRALIZADO Or vFld(Rs("Estado")) = ED_PAGADO Then
                     Q1 = Q1 & ", Estado = " & ED_PAGADO
                     SetPagado = True
                  End If
                  
               ElseIf Not SetPagado Then    'Saldo = Total
                  
                  If (vFld(Rs("IdCompCent")) <> 0 Or Year(vFld(Rs("FEmision"))) < Ano) Then    'está asociado a un comprobante de centralización o es del año anterior
                     Q1 = Q1 & ", Estado = " & ED_CENTRALIZADO
                  Else
                     Q1 = Q1 & ", Estado = " & ED_PENDIENTE
                  End If
               
               End If
            End If
            
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
            Call ExecSQL(vDbMain, Q1)
         End If
         
      End If
         
      CurIdDoc = vFld(Rs("IdDoc"))
      ' 2828725 Cambia a estado pendiente los documentos NDV si los comprobantes no suman igual el debe y el haber
'        Q1 = "SELECT Switch(sum(debe) = sum(haber), 0,sum(debe) <> sum(haber),  abs(sum(debe) - sum(haber))) as pagado "
'        Q1 = Q1 & " FROM    documento as docu, movcomprobante as mov, comprobante com "
'        Q1 = Q1 & " WHERE   docu.iddoc = mov.iddoc "
'        Q1 = Q1 & " AND     mov.idcomp = com.idcomp "
'        Q1 = Q1 & " AND tipolib = 2 AND tipodoc = 4 "
'        Q1 = Q1 & " AND     docu.numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'        Set Rs2 = OpenRs(vDbMain, Q1)
'
'
'        Do While Rs2.EOF = False
'
'            If vFld(Rs2("pagado")) > 0 Then
'
'            Q1 = "UPDATE documento "
'            Q1 = Q1 & " SET Estado = " & ED_PENDIENTE
'            Q1 = Q1 & " , SaldoDoc = " & vFld(Rs2("pagado"))
'            Q1 = Q1 & " WHERE numdoc = '" & vFld(Rs("IdDoc")) & "' "
'
'            Call ExecSQL(vDbMain, Q1)
'
'            End If
'        Rs2.MoveNext
'        Loop
'        Call CloseRs(Rs2)
        ' FIN 2828725
               
      
      Rs.MoveNext
   Loop
      
   Call CloseRs(Rs)
   
     
End Function 'genera los docs que quedaron pendientes del año anterior
'14520904

'14520904
Public Sub CorregirSaldosAnoAnterior()

   Dim WhLib As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim n As Long
   Dim PathDbAnoAnt As String
   Dim FechaBase As Long
   Dim Rs1 As Recordset
   
#If DATACON = 1 Then
Dim DbAnoAnt As Database
#Else
Dim DbAnoAnt As ADODB.Connection
Set DbAnoAnt = DbMain
#End If
   
   'Dim DbAnoAnt As Database
   Dim ConnStr As String
   'Dim Q1 As String
   Dim vSaldos As String
   
   vSaldos = "0"
   
   Q1 = ""
   Q1 = Q1 & " SELECT Codigo From ParamEmpresa where Tipo = 'SALDOS' " ' and Valor ='" & ParaSQL(W.Version) & "'"
   Q1 = Q1 & " and  IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " and  Ano = " & gEmpresa.Ano
   
  Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      vSaldos = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
 
   
If vSaldos = "0" Then
   
    If gDbType = SQL_ACCESS Then
    
     PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")
            
     If ExistFile(PathDbAnoAnt) Then
         
         'cerramos el año actual y abrimos el año anterior
       Call CloseDb(DbMain)
       
       Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
      ' Call LinkMdbAdm
       
       'corrige base del año anterior, por si las moscas
       #If DATACON = 1 Then
       Call CorrigeBase
       #End If
       
       'cerramos el año anterior y abrimos el año actual
       Call CloseDb(DbMain)
       Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
     Else
          Exit Sub
     End If



    PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")
        
    If ExistFile(PathDbAnoAnt) Then
      ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
      Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
      
    Else
      Exit Sub
    End If
End If
   
   WhLib = " Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") "

   Set Rs = OpenRs(DbAnoAnt, "SELECT Count(*) FROM Documento WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1)
   If Not Rs.EOF Then
      n = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
   'Me.MousePointer = vbHourglass
   
   'asignamos SaldoDoc = NULL para los docs de compras, ventas y retenciones para que los recalcule TOTOS
   Call ExecSQL(DbAnoAnt, "UPDATE Documento SET SaldoDoc = NULL WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1)
   
   Call RecalcSaldosAnoAnterior(DbAnoAnt, gEmpresa.id, gEmpresa.Ano - 1)
   'Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
   'Call RecalcSaldosFulle(gEmpresa.id, gEmpresa.Ano - 1)
   
  ' Call CloseDb(DbAnoAnt)
  Call GenDocsPendientes(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, False, True)
  
  
   'Me.MousePointer = vbDefault
  Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
  'Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
   
  Q1 = "INSERT INTO ParamEmpresa (IdEmpresa,Ano,Tipo, Codigo, Valor ) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'SALDOS','1' , '" & ParaSQL(W.Version) & "' )"
   
  Call ExecSQL(DbMain, Q1)
    
   'Prueba Mejora para los documentos que cambian el estado en el año siguiente
  FechaBase = DateSerial(gEmpresa.Ano - 1, 12, 31)
   Q1 = ""
   Q1 = Q1 & " SELECT NumDoc, IdEmpresa, Ano, TipoLib, TipoDoc, IdEntidad, IdDoc From Documento "
   Q1 = Q1 & " Where  IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " And  FEmision <=  " & FechaBase
   Q1 = Q1 & " And  Estado =  " & ED_PENDIENTE

   Set Rs = OpenRs(DbMain, Q1)
   Do While Rs.EOF = False
 
        Q1 = ""
        Q1 = Q1 & " SELECT Estado From Documento "
        Q1 = Q1 & " Where NumDoc = '" & vFld(Rs("NumDoc")) & "'"
        Q1 = Q1 & " And IdEmpresa = " & vFld(Rs("IdEmpresa"))
        Q1 = Q1 & " And Ano = " & vFld(Rs("Ano")) - 1
        Q1 = Q1 & " And TipoLib = " & vFld(Rs("TipoLib"))
        Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
        Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
        Set Rs1 = OpenRs(DbAnoAnt, Q1)
        If Not Rs1.EOF Then
            If vFld(Rs1("Estado")) = ED_CENTRALIZADO Or vFld(Rs1("Estado")) = ED_PAGADO Then
                Call ExecSQL(DbMain, "UPDATE Documento SET Estado = " & vFld(Rs1("Estado")) & " WHERE IdDoc = " & vFld(Rs("IdDoc")))
            End If
        End If
        Call CloseRs(Rs1)
   Rs.MoveNext
   Loop
   Call CloseRs(Rs)
    
 End If
        
 #If DATACON <> 1 Then
 '14704978
    Call CorrigeCuentaCompTipo
'14704978
   #End If
   If gDbType = SQL_ACCESS Then
   
   'Call CorrigeDuplicados(False)
'   Call CorrigePagadosAñoAnteriores(False)
  
'   Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
'   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)

     'Call CorrigePagadosAñoAnteriores2(False)
     Call CorrigePagadosAñoAnterioresPrueba(False)
     
'    If CorrigeDuplicados(False) = True Then
'       Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
'       Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
'    End If

   Else
     'Call CorrigePagadosAñoAnteriores2(False)
     Call CorrigePagadosAñoAnterioresPrueba(False)
     
     
   End If
End Sub

Public Sub limpiarCampoTotPagado()
Dim Q1 As String
 '3125609
                 Q1 = "UPDATE Documento SET TotPagadoAnoAnt = 0 "
                 Q1 = Q1 & " Where IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                 Q1 = Q1 & " AND Year(FEmision) = " & gEmpresa.Ano
                 Q1 = Q1 & " AND (FExported = 0 Or FExported is null) "
                 Q1 = Q1 & " AND TotPagadoAnoAnt <> null "
                                  
                 Call ExecSQL(DbMain, Q1)
                '3125609

End Sub

''SF 14202137
Public Function CorrigeDuplicados(ByVal vMsj As Boolean) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim Rs4 As Recordset
   Dim Rs5 As Recordset
   Dim Rs6 As Recordset
   Dim Rs7 As Recordset
   Dim vTiene As Boolean

   On Error Resume Next

    ERR.Clear
    
    CorrigeDuplicados = False
    
Q1 = ""
Q1 = Q1 & " SELECT NUMDOC, TIPODOC, IDENTIDAD, IDEMPRESA, ANO, COUNT(NUMDOC) as Cant FROM DOCUMENTO "

Q1 = Q1 & " WHERE TIPODOC not in (19,20,15) " ' 3160844 se saca los tipo doc VPE,VPEE y VSD

Q1 = Q1 & " AND IDEMPRESA = " & gEmpresa.id
'se agrega tipo lib otros doc y retenciones para no entrar al proceso
Q1 = Q1 & " AND TipoLib not in (5,4) "
Q1 = Q1 & " GROUP BY NUMDOC, TIPODOC, IDENTIDAD, IDEMPRESA, ANO "
Q1 = Q1 & " Having Count(NumDoc) > 10 "
Q1 = Q1 & " ORDER BY 3 DESC"

Set Rs = OpenRs(DbMain, Q1)

Do While Not Rs.EOF

CorrigeDuplicados = True

' Q2 = ""
' Q2 = " SELECT IDDOC FROM DOCUMENTO WHERE NUMDOC = '" & vFld(Rs("NUMDOC")) & "'"
' Set Rs2 = OpenRs(DbMain, Q2)
' Do While Not Rs2.EOF
' Sleep 1000

Q3 = ""
Q3 = "delete from movDocumento "
Q3 = Q3 & " where IDDOC < 0 " ' = " & vFld(Rs2("IDDOC"))
Call ExecSQL(DbMain, Q3)

Q3 = ""
Q3 = "delete from documento "
Q3 = Q3 & " where NUMDOC = '" & vFld(Rs("NUMDOC")) & "'"
Call ExecSQL(DbMain, Q3)

' Rs2.MoveNext
'
' Loop
' Call CloseRs(Rs2)
Rs.MoveNext

Loop

Call CloseRs(Rs)
    
    Q1 = ""
    Q1 = Q1 & " SELECT NUMDOC, TIPODOC,TipoLib, IDENTIDAD, COUNT(NUMDOC) as Cant FROM DOCUMENTO "
    Q1 = Q1 & " WHERE TIPODOC not in (19,20,15) " ' 3160844 se saca los tipo doc VPE,VPEE y VSD
    Q1 = Q1 & " AND IDEMPRESA = " & gEmpresa.id
    'se agrega tipo lib otros doc y retenciones para no entrar al proceso
    Q1 = Q1 & " AND TipoLib not in (5,4) "
    Q1 = Q1 & " GROUP BY NUMDOC, TIPODOC, IDENTIDAD,TipoLib "
    Q1 = Q1 & " Having Count(NumDoc) > 1 "
    Q1 = Q1 & " ORDER BY 3 DESC"

    Set Rs = OpenRs(DbMain, Q1)

    Do While Not Rs.EOF
        
      CorrigeDuplicados = True
        
'        If vFld(Rs("TipoLib")) = LIB_VENTAS And vFld(Rs("TIPODOC")) = TDOC_BOLVENTA Then
'
'        End If
        
        vTiene = False
            
        Q2 = ""
        Q2 = Q2 & "SELECT DO.IDDOC, MOV.IDCOMP,iif(mov.iddoc is null,0,1) as tiene "
        Q2 = Q2 & " FROM (DOCUMENTO AS DO LEFT JOIN MOVCOMPROBANTE AS MOV ON MOV.IDDOC = DO.IDDOC)"
        Q2 = Q2 & " WHERE NUMDOC = '" & vFld(Rs("NUMDOC")) & "'"
'        Q2 = Q2 & " AND NUMDOC = '" & vFld(Rs("NUMDOC")) & "'"
'        Q2 = Q2 & " AND NUMDOC = '" & vFld(Rs("NUMDOC")) & "'"
        Q2 = Q2 & " GROUP BY DO.IDDOC, MOV.IDCOMP,iif(mov.iddoc is null,0,1)"
        Q2 = Q2 & " ORDER BY 3 DESC "
        
         Set Rs2 = OpenRs(DbMain, Q2)

         Do While Not Rs2.EOF
         
         If vFld(Rs2("tiene")) = 1 Then
          vTiene = True
         End If
         
'          If vFld(Rs("NUMDOC")) = 107594 Then
'         MsgBox1 ""
'        End If
         
    If vTiene Then
         
         If (Rs2("tiene")) = 1 Then
                      
                  Q2 = ""
                  Q2 = Q2 & "Select count(*) as centra  from comprobante "
                  Q2 = Q2 & " where idcomp = " & vFld(Rs2("IDCOMP"))
                  Q2 = Q2 & " and tipo = " & TC_TRASPASO
            
                  Set Rs3 = OpenRs(DbMain, Q2)
                  
                  If Rs3.BOF = False Then
                    
                    If vFld(Rs3("centra")) = 1 Then
                      
                      Q3 = "UPDATE documento SET estado =  " & ED_CENTRALIZADO
                      Q3 = Q3 & ", IdCompCent = " & vFld(Rs2("IDCOMP"))
                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                      Call ExecSQL(DbMain, Q3)
                    End If
                    
                  End If
                  
                  Call CloseRs(Rs3)
              
           Else
                  Q3 = ""
                  Q3 = "delete from movDocumento "
                  Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                  Call ExecSQL(DbMain, Q3)
                 
                  Q3 = ""
                  Q3 = "delete from documento "
                  Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                  Call ExecSQL(DbMain, Q3)
          End If
      Else
      
             If vFld(Rs("Cant")) > 2 Then
                Q2 = ""
                Q2 = Q2 & "SELECT count(idDoc) as Cant from documento "
                Q2 = Q2 & " WHERE IdDoc = " & vFld(Rs2("iddoc"))
                Q2 = Q2 & " and FExported = 0 or FExported is null"
                
                 Set Rs4 = OpenRs(DbMain, Q2)
        
              If Not Rs4.EOF Then
                  If vFld(Rs4("Cant")) > 0 Then
                    
                      Q3 = ""
                      Q3 = "delete from movDocumento "
                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                      Call ExecSQL(DbMain, Q3)
                       
                      Q3 = ""
                      Q3 = "delete from documento "
                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                      Call ExecSQL(DbMain, Q3)
                      End If
                    
                 End If
                 
              Call CloseRs(Rs4)
             Else
             
                Q2 = ""
                Q2 = Q2 & "SELECT count(idDoc) as Cant from documento "
                Q2 = Q2 & " WHERE IdDoc = " & vFld(Rs2("iddoc"))
                'Q2 = Q2 & " and (OldIdDocTmp > 0 or OldIdDocTmp is null) "
                Q2 = Q2 & " and (OldIdDocTmp = 0 or OldIdDocTmp is null) "
                
                 Set Rs5 = OpenRs(DbMain, Q2)
        
                 If Not Rs5.EOF Then
                      If vFld(Rs5("Cant")) > 0 Then
                    
                        Q3 = ""
                        Q3 = "delete from movDocumento "
                        Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                        Call ExecSQL(DbMain, Q3)
                       
                        Q3 = ""
                        Q3 = "delete from documento "
                        Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                        Call ExecSQL(DbMain, Q3)
                         
                       Else
                 
                  '3149389
                  
                       Dim PathDbAnoAnt As String
                     
                       Dim DbAnoAnt As Database
                       Dim ConnStr As String
                    
                    
                    If gDbType = SQL_ACCESS Then
                        PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")
                            
                        If ExistFile(PathDbAnoAnt) Then
                          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
                          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
                          
                        Else
                          Exit Function
                        End If
                    End If
                  
                   
                    Q1 = ""
                    Q1 = Q1 & " SELECT IdDoc,NUMDOC FROM DOCUMENTO "
                    Q1 = Q1 & " Where NumDoc = '" & vFld(Rs("NumDoc")) & "'"
                    Q1 = Q1 & " And TIPODOC =  " & vFld(Rs("TIPODOC"))
                    Q1 = Q1 & " And IDENTIDAD = " & vFld(Rs("IDENTIDAD"))
                    Q1 = Q1 & " And TipoLib = " & vFld(Rs("TipoLib"))
                
                    Set Rs6 = OpenRs(DbAnoAnt, Q1)
                
                   If Not Rs6.EOF Then
                      
                        Q1 = ""
                        Q1 = Q1 & " SELECT IdDoc,NUMDOC FROM DOCUMENTO "
                        Q1 = Q1 & " Where NumDoc = '" & vFld(Rs("NumDoc")) & "'"
                        Q1 = Q1 & " And TIPODOC =  " & vFld(Rs("TIPODOC"))
                        Q1 = Q1 & " And IDENTIDAD = " & vFld(Rs("IDENTIDAD"))
                        Q1 = Q1 & " And TipoLib = " & vFld(Rs("TipoLib"))
                        Q1 = Q1 & " And OldIdDoc <> " & vFld(Rs6("IdDoc"))
                        
                        Set Rs7 = OpenRs(DbMain, Q1)
                        
                          If Not Rs7.EOF Then
                           
                              Q3 = ""
                              Q3 = "delete from movDocumento "
                              Q3 = Q3 & " where iddoc = " & vFld(Rs7("iddoc"))
                              Call ExecSQL(DbMain, Q3)
        
                              Q3 = ""
                              Q3 = "delete from documento "
                              Q3 = Q3 & " where iddoc = " & vFld(Rs7("iddoc"))
                              Call ExecSQL(DbMain, Q3)
                          
                          End If
                        Call CloseRs(Rs7)
                    End If
                    
                     Call CloseRs(Rs6)
                    
                    Call CloseDb(DbAnoAnt)
                   '3149389
                  
                  End If
                                
              End If
                
                 Call CloseRs(Rs5)
             End If
      End If
         
         Rs2.MoveNext
         
         'CorrigeDuplicados = True
         Loop
        
       Call CloseRs(Rs2)

         Rs.MoveNext

      Loop

      Call CloseRs(Rs)
      
   If vMsj Then
  MsgBox1 "Documentos duplicados eliminados correctamente.", vbInformation + vbOKOnly
   End If
End Function
''SF 14202137

''SF 14202137
Public Sub CorrigePagadosAñoAnteriores(ByVal vMsj As Boolean)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim Rs4 As Recordset
   Dim Rs5 As Recordset
   
   Dim i As Integer
   Dim AñoEliminacion As Long
   
   
  ' Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
   
   On Error Resume Next
   
        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If

   

    ERR.Clear
    
    Dim FechaEmiOriginal As Long, FechaLong As Long
    FechaEmiOriginal = DateSerial(gEmpresa.Ano, GetMesActual() + 1, 1 - 1)
    
    Q1 = ""
    Q1 = Q1 & " SELECT Documento.NumDoc, Documento.TipoLib,  Documento.TipoDoc, Documento.IdEntidad ,FEmision,Documento.IdEmpresa "
    Q1 = Q1 & " FROM (((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad  )  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc) "
    Q1 = Q1 & " LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta) "
    Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc "
    Q1 = Q1 & " WHERE  (EsTotalDoc <> 0 OR EsTotalDoc IS NULL) AND Documento.IdEntidad > 0 AND   (Documento.TipoLib  IN(1,2,3)  OR (Documento.TipoLib IN (5,4) AND Documento.DocOtrosEnAnalitico <> 0))"
    Q1 = Q1 & " AND Documento.Estado <> 1  AND (vMovCompIdDoc.IdDoc IS NULL   OR (Documento.Estado IN (3,4) AND Documento.IdCompCent = 0 )) " ' AND MovDocumento.IdCuenta = 372 "
    Q1 = Q1 & " AND Documento.FEmisionOri <= " & CLng(FechaEmiOriginal) & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
    If gDbType = SQL_ACCESS Then
    Q1 = Q1 & " and year(Documento.FEmision) < " & gEmpresa.Ano & " and documento.estado in (" & ED_CENTRALIZADO & "," & ED_PAGADO & ")"
    Else
    Q1 = Q1 & " and year(Documento.FEmision -2) < " & gEmpresa.Ano & " and documento.estado in (" & ED_CENTRALIZADO & "," & ED_PENDIENTE & "," & ED_PAGADO & ")"
    End If
    
    Q1 = Q1 & " GROUP BY Documento.NumDoc, Documento.TipoLib,   Documento.TipoDoc, Documento.IdEntidad,Documento.FEmision,Documento.IdEmpresa"
    
    Set Rs = OpenRs(DbMain, Q1)

    Do While Not Rs.EOF
        
        Dim fechaemi As String
        
        fechaemi = Format(vFld(Rs("FEmision")), SDATEFMT)
        
         If gDbType = SQL_ACCESS Then
             'PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & Year(fechaemi) & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 
                 PathDbAnoAnt = Replace(PathDbAnoAnt, "\..", "")
             
             If ExistFile(PathDbAnoAnt) Then
               ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
               Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
               
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
               Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
               Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
               Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
               Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
               
               Set Rs2 = OpenRs(DbAnoAnt, Q1)
               
               If Rs2.EOF = False Then
               
                  If vFld(Rs2("Estado")) = ED_PAGADO And vFld(Rs2("SaldoDoc")) = 0 Then
                         
                        
                        Do While Year(fechaemi) <= gEmpresa.Ano
                            
                             PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & Year(fechaemi) & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 
                             PathDbAnoAnt = Replace(PathDbAnoAnt, "\..", "")
                             
                             If ExistFile(PathDbAnoAnt) Then
                               ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
                               Call CloseDb(DbAnoAnt)
                               Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
                                 
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
                                 Q1 = Q1 & "From documento "
                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                 Q1 = Q1 & " And estado <> " & ED_PAGADO
                                 Set Rs3 = OpenRs(DbAnoAnt, Q1)
                                 
                                 If Rs3.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbAnoAnt, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbAnoAnt, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs3)
                               
                             End If
                             
                             fechaemi = DateAdd("YYYY", 1, fechaemi)
                             
                         Loop
                    End If
                    
               End If
               Call CloseRs(Rs2)
             End If
      Else
      
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
               Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
               Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
               Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
               Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
               Q1 = Q1 & " and idempresa = " & vFld(Rs("IdEmpresa"))
               Q1 = Q1 & " and ano = " & gEmpresa.Ano - 1
               
'               If vFld(Rs("NumDoc")) = 3008 Then
'                MsgBox ""
'               End If
               
               Set Rs2 = OpenRs(DbAnoAnt, Q1)
               
               If Rs2.EOF = False Then
               
                  If vFld(Rs2("Estado")) = ED_PAGADO And vFld(Rs2("SaldoDoc")) = 0 Then
                         
                       ' Dim x As Integer
                        
                        'Do While Year(fechaemi) <= gEmpresa.Ano
                                                             
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
                                 Q1 = Q1 & "From documento "
                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                ' Q1 = Q1 & " And estado <> " & ED_PAGADO
                                 Q1 = Q1 & " And idEmpresa =" & vFld(Rs("Idempresa"))
                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
                                 
                                 Set Rs3 = OpenRs(DbAnoAnt, Q1)
                                 
                                 If Rs3.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      'Q3 = Q3 & " And "
                                      Call ExecSQL(DbAnoAnt, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbAnoAnt, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs3)
                               
                            ' End If
                             
                             fechaemi = DateAdd("YYYY", 1, fechaemi)
                             
                         'Loop
                    End If
                Else
                
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
                                 Q1 = Q1 & "From documento "
                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                ' Q1 = Q1 & " And estado <> " & ED_PAGADO
                                 Q1 = Q1 & " And idEmpresa =" & vFld(Rs("Idempresa"))
                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
                                 
                                 Set Rs3 = OpenRs(DbAnoAnt, Q1)
                                 
                                 If Rs3.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      'Q3 = Q3 & " And "
                                      Call ExecSQL(DbAnoAnt, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbAnoAnt, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs3)
                    
               End If
               Call CloseRs(Rs2)
             End If
             
      
        
    Rs.MoveNext

    Loop

      Call CloseRs(Rs)

   If vMsj Then
  MsgBox1 "Documentos pagados años anteriores eliminados correctamente.", vbInformation + vbOKOnly
   End If
End Sub
''SF 14202137

''SF 14202137
Public Sub CorrigePagadosAñoAnteriores2(ByVal vMsj As Boolean)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim Rs4 As Recordset
   Dim Rs5 As Recordset
   
   Dim i As Integer
   Dim AñoEliminacion As Long
   
   
  ' Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
   
   On Error Resume Next
   
   
        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If
    ERR.Clear
    
    Dim FechaEmiOriginal As Long, FechaLong As Long
    FechaEmiOriginal = DateSerial(gEmpresa.Ano, GetMesActual() + 1, 1 - 1)
    
    Q1 = ""
    Q1 = Q1 & " SELECT Documento.NumDoc, Documento.TipoLib,  Documento.TipoDoc, Documento.IdEntidad ,FEmision,Documento.IdEmpresa "
    Q1 = Q1 & " FROM (((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad  )  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc) "
    Q1 = Q1 & " LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta) "
    Q1 = Q1 & " LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc "
    Q1 = Q1 & " WHERE  (EsTotalDoc <> 0 OR EsTotalDoc IS NULL) AND Documento.IdEntidad > 0 AND   (Documento.TipoLib  IN(1,2,3)  OR (Documento.TipoLib IN (5,4) AND Documento.DocOtrosEnAnalitico <> 0))"
    Q1 = Q1 & " AND Documento.Estado <> 1  AND (vMovCompIdDoc.IdDoc IS NULL   OR (Documento.Estado IN (3,4) AND Documento.IdCompCent = 0 )) " ' AND MovDocumento.IdCuenta = 372 "
    Q1 = Q1 & " AND Documento.FEmisionOri <= " & CLng(FechaEmiOriginal) & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
    If gDbType = SQL_ACCESS Then
    Q1 = Q1 & " and year(Documento.FEmision) < " & gEmpresa.Ano & " and documento.estado in (" & ED_CENTRALIZADO & "," & ED_PAGADO & ")"
    Else
    Q1 = Q1 & " and year(Documento.FEmision -2) < " & gEmpresa.Ano & " and documento.estado in (" & ED_CENTRALIZADO & "," & ED_PENDIENTE & "," & ED_PAGADO & ")"
    End If
    
    Q1 = Q1 & " GROUP BY Documento.NumDoc, Documento.TipoLib,   Documento.TipoDoc, Documento.IdEntidad,Documento.FEmision,Documento.IdEmpresa"
    
    Set Rs = OpenRs(DbMain, Q1)

    Do While Not Rs.EOF
    
        
        Dim fechaemi As String
        
        fechaemi = Format(vFld(Rs("FEmision")), SDATEFMT)
        
         If gDbType = SQL_ACCESS Then
             PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 'PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & Year(fechaemi) & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 
                 PathDbAnoAnt = Replace(PathDbAnoAnt, "\..", "")
             
             If ExistFile(PathDbAnoAnt) Then
               ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
               Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
               
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
               Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
               Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
               Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
               Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
               
               Set Rs2 = OpenRs(DbAnoAnt, Q1)
               
               If Rs2.EOF = False Then
               
                  If vFld(Rs2("Estado")) = ED_PAGADO And vFld(Rs2("SaldoDoc")) = 0 Then
                         
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
                                 Q1 = Q1 & "From documento "
                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                 'Q1 = Q1 & " And estado <> " & ED_PAGADO
                                 Set Rs3 = OpenRs(DbMain, Q1)
                                 
                                 If Rs3.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbMain, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbMain, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs3)
                               
                             End If
                             
                             fechaemi = DateAdd("YYYY", 1, fechaemi)
                                        
                    'End If
                Else
                
'                               Q1 = ""
'                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
'                                 Q1 = Q1 & "From documento "
'                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
'                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
'                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
'                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
'                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
'                                 'Q1 = Q1 & " And estado <> " & ED_PAGADO
'                                 Set Rs3 = OpenRs(DbMain, Q1)
'
'                                 If Rs3.EOF = False Then
'
'                                      Q3 = ""
'                                      Q3 = "delete from movDocumento "
'                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
'                                      Call ExecSQL(DbMain, Q3)
'
'                                      Q3 = ""
'                                      Q3 = "delete from documento "
'                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
'                                      Call ExecSQL(DbMain, Q3)
'                                  End If
'
'                                  Call CloseRs(Rs3)

               End If
               Call CloseRs(Rs2)
             End If
             
           Call CloseDb(DbAnoAnt)
      Else
      
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
               Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
               Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
               Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
               Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
               Q1 = Q1 & " and idempresa = " & vFld(Rs("IdEmpresa"))
               Q1 = Q1 & " and ano = " & gEmpresa.Ano - 1
               
'               If vFld(Rs("NumDoc")) = 3008 Then
'                MsgBox ""
'               End If
               
               Set Rs2 = OpenRs(DbAnoAnt, Q1)
               
               If Rs2.EOF = False Then
               
                  If vFld(Rs2("Estado")) = ED_PAGADO And vFld(Rs2("SaldoDoc")) = 0 Then
                         
                       ' Dim x As Integer
                        
                        'Do While Year(fechaemi) <= gEmpresa.Ano
                                                             
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
                                 Q1 = Q1 & "From documento "
                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                ' Q1 = Q1 & " And estado <> " & ED_PAGADO
                                 Q1 = Q1 & " And idEmpresa =" & vFld(Rs("Idempresa"))
                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
                                 
                                 Set Rs3 = OpenRs(DbAnoAnt, Q1)
                                 
                                 If Rs3.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      'Q3 = Q3 & " And "
                                      Call ExecSQL(DbAnoAnt, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
                                      Call ExecSQL(DbAnoAnt, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs3)
                               
                            ' End If
                             
                             fechaemi = DateAdd("YYYY", 1, fechaemi)
                             
                         'Loop
                    End If
                Else
                
'                                 Q1 = ""
'                                 Q1 = Q1 & "SELECT IdDoc,Estado,SaldoDoc "
'                                 Q1 = Q1 & "From documento "
'                                 Q1 = Q1 & "Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
'                                 Q1 = Q1 & "And TipoLib = " & vFld(Rs("TipoLib"))
'                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
'                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
'                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
'                                ' Q1 = Q1 & " And estado <> " & ED_PAGADO
'                                 Q1 = Q1 & " And idEmpresa =" & vFld(Rs("Idempresa"))
'                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
'
'                                 Set Rs3 = OpenRs(DbAnoAnt, Q1)
'
'                                 If Rs3.EOF = False Then
'
'                                      Q3 = ""
'                                      Q3 = "delete from movDocumento "
'                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
'                                      'Q3 = Q3 & " And "
'                                      Call ExecSQL(DbAnoAnt, Q3)
'
'                                      Q3 = ""
'                                      Q3 = "delete from documento "
'                                      Q3 = Q3 & " where iddoc = " & vFld(Rs3("iddoc"))
'                                      Call ExecSQL(DbAnoAnt, Q3)
'                                  End If
'
'                                  Call CloseRs(Rs3)

               End If
               Call CloseRs(Rs2)
             End If
        
    Rs.MoveNext

    Loop

      Call CloseRs(Rs)

   If vMsj Then
  MsgBox1 "Documentos pagados años anteriores eliminados correctamente.", vbInformation + vbOKOnly
   End If
End Sub
''SF 14202137


'3126513
#If DATACON <> 1 Then
Public Sub CorrigeCuentaCompTipo()
Dim Q1 As String

  '3133472

   Q1 = ""
   Q1 = Q1 & "if not exists (select name from sysindexes  where name = 'idx_Comp_ComTipo') "
   Q1 = Q1 & "CREATE NONCLUSTERED INDEX idx_Comp_ComTipo ON Comprobante "
   Q1 = Q1 & "(IdComp ASC,IdEmpresa ASC,Ano ASC); "
   
   Call ExecSQL(DbMain, Q1)
   
   Q1 = ""
   Q1 = Q1 & "if not exists (select name from sysindexes  where name = 'idx_Cuentas_ComTipo') "
   Q1 = Q1 & "CREATE NONCLUSTERED INDEX idx_Cuentas_ComTipo ON Cuentas "
   Q1 = Q1 & "(IdEmpresa ASC,Ano ASC,idCuenta ASC,Codigo ASC); "
   
   Call ExecSQL(DbMain, Q1)
   
   Q1 = ""
   Q1 = Q1 & "if not exists (select name from sysindexes  where name = 'idx_MovComprobante_ComTipo') "
   Q1 = Q1 & "CREATE NONCLUSTERED INDEX idx_MovComprobante_ComTipo ON MovComprobante "
   Q1 = Q1 & "(IdEmpresa ASC,Ano ASC,IdComp ASC,IdCuenta ASC); "
   
   Call ExecSQL(DbMain, Q1)
   '3133472

    Q1 = ""
    Q1 = Q1 & "Update MovComprobante "
    Q1 = Q1 & "Set IdCuenta = t.cta2022 from ( "
    Q1 = Q1 & "Select d.IdComp as idcompr, d.IdCuenta cta2021, "
    Q1 = Q1 & "(Select p.idCuenta "
    Q1 = Q1 & "From cuentas as p where p.Codigo = d.codigo "
    Q1 = Q1 & "And p.Ano = " & gEmpresa.Ano - 1
    Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id & ") cta2022 "
    Q1 = Q1 & "From ( "
    Q1 = Q1 & "Select Comprobante.IdComp,correlativo,Comprobante.Ano,MovComprobante.IdCuenta,Comprobante.IdEmpresa, Cuentas.Codigo "
    Q1 = Q1 & "From Comprobante, MovComprobante, Cuentas "
    Q1 = Q1 & "Where Comprobante.IdComp = MovComprobante.IdComp "
    Q1 = Q1 & "And Comprobante.Ano = MovComprobante.Ano "
    Q1 = Q1 & "And MovComprobante.IdCuenta = Cuentas.idCuenta "
    Q1 = Q1 & "AND Comprobante.IdEmpresa = Cuentas.IdEmpresa "
    Q1 = Q1 & "And Cuentas.Ano = " & gEmpresa.Ano - 2
    Q1 = Q1 & " And Comprobante.Ano =" & gEmpresa.Ano - 1
    Q1 = Q1 & " AND Comprobante.IdEmpresa =" & gEmpresa.id
    Q1 = Q1 & " ) as d) as t "
    Q1 = Q1 & "Where IdComp = t.idcompr And IdCuenta = t.cta2021 "
    Q1 = Q1 & "And Ano = " & gEmpresa.Ano - 1
    Q1 = Q1 & " and IdEmpresa = " & gEmpresa.id
           
    Call ExecSQL(DbMain, Q1)
       
    Q1 = ""
    Q1 = Q1 & "Update MovComprobante "
    Q1 = Q1 & "Set IdCuenta = t.cta2022 from ( "
    Q1 = Q1 & "Select d.IdComp as idcompr, d.IdCuenta cta2021, "
    Q1 = Q1 & "(Select p.idCuenta "
    Q1 = Q1 & "From cuentas as p where p.Codigo = d.codigo "
    Q1 = Q1 & "And p.Ano = " & gEmpresa.Ano
    Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id & ") cta2022 "
    Q1 = Q1 & "From ( "
    Q1 = Q1 & "Select Comprobante.IdComp,correlativo,Comprobante.Ano,MovComprobante.IdCuenta,Comprobante.IdEmpresa, Cuentas.Codigo "
    Q1 = Q1 & "From Comprobante, MovComprobante, Cuentas "
    Q1 = Q1 & "Where Comprobante.IdComp = MovComprobante.IdComp "
    Q1 = Q1 & "And Comprobante.Ano = MovComprobante.Ano "
    Q1 = Q1 & "And MovComprobante.IdCuenta = Cuentas.idCuenta "
    Q1 = Q1 & "AND Comprobante.IdEmpresa = Cuentas.IdEmpresa "
    Q1 = Q1 & "And Cuentas.Ano = " & gEmpresa.Ano - 1
    Q1 = Q1 & " And Comprobante.Ano =" & gEmpresa.Ano
    Q1 = Q1 & " AND Comprobante.IdEmpresa =" & gEmpresa.id
    Q1 = Q1 & " ) as d) as t "
    Q1 = Q1 & "Where IdComp = t.idcompr And IdCuenta = t.cta2021 "
    Q1 = Q1 & "And Ano = " & gEmpresa.Ano
    Q1 = Q1 & " and IdEmpresa = " & gEmpresa.id
               
    Call ExecSQL(DbMain, Q1)
    
  
End Sub
 #End If
'3126513
 
'Duplicados 14690904
Public Function ComprobanteApeturaFexported()
Dim Rs As Recordset
   Dim Q1 As String
   Dim InitAno As String
   Dim IdCompAper As Long
   Dim Rc As Integer
   Dim HayComp As Boolean
   Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
    
'   If Not ValidaIngresoComp(True) Then
'      Exit Function
'   End If
      
   
      
   'veamos si la empresa tiene historia (año anterior a partir del cual se generó este año)
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='INITAÑO' "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      InitAno = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
         
   'veamos si ya se generó comprobante de apertura
   Q1 = "SELECT IdCompAper"
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      IdCompAper = vFld(Rs("IdCompAper"))
   End If
   
   Call CloseRs(Rs)
   
   If IdCompAper = 0 Then
   
      ' Veamos si exsite un comprobante de aprtura generado manualmente, dado que no está registrado en la tabla EmpresasAño
      Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdCompAper = vFld(Rs("IdComp"))
      End If
      
      Call CloseRs(Rs)
   
   End If
   
  
   'si ya existe un comp de apertura, lo regeneramos
   If IdCompAper <> 0 Then

      'veamos si el ID del comprobante deapertura corresponde al almacenado en la tabla EmpresasAño de la LPContab
      Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         
         If IdCompAper <> vFld(Rs("IdComp")) Then   'no calza el IdCompAper de EmpresasAño con el comprobante en la base del año
            
            IdCompAper = 0
            
            Q1 = "UPDATE EmpresasAno SET IdCompAper = 0 "
            Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      Else
         IdCompAper = 0
                  
      End If
      
      Call CloseRs(Rs)
      
      Rc = GenCompAperturaDuplicados(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
   
   Else        'IdCompAper = 0
           
      Q1 = "SELECT IdComp, Fecha FROM Comprobante "
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY Fecha DESC, IdComp DESC"
      Set Rs = OpenRs(DbMain, Q1)
      
      HayComp = (Not Rs.EOF)
      
      Call CloseRs(Rs)
   
      If Not HayComp Then 'no hay comprobantes aún
         Rc = GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
      
      ElseIf gTipoCorrComp = TCC_TIPOCOMP Then 'hay comprobantes pero el correlativo es por tipo => podemos generar un comp de apertura con N° 1
         Rc = GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
      
      Else
      '   MsgBox1 "No es posible generar automáticamente el comprobante de apertura, dado que ya hay comprobantes ingresados y éste debe ser el primero.", vbExclamation
         'Me.MousePointer = vbDefault
         Exit Function
      End If
      
   End If
   
  ' Me.MousePointer = vbDefault
      
   If Rc Then
     ' MsgBox1 "El Comprobante de apertura ha sido generado.", vbInformation + vbOKOnly
      
     ' Me.MousePointer = vbHourglass
      
      If InitAno = "EMPHISTORIA" Then
         
         
         Call GenDocsPendientes(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, False, True)
         'Call GenDocsFullPendientes(gEmpresa.Id, gEmpresa.Rut, gEmpresa.Ano, True)
         
         
         'Call GenActFijoResidual(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True)
      ElseIf InitAno = "EMPHISTACC" Then
        
        PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
        
        If ExistFile(PathDbAnoAnt) Then
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
         
          'ahora obtenemos los documentos centralizados y pagados con saldo pendiente desde el año anterior
          
          Call CopyDocsFromAccessToSQLServerNew(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
          
          'Luego los activos fijos con valor libro mayor que cero o no depreciables del año anteriro
          
          'Call CopyActFijoFromAccessToSQLServer(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
          
          'finalmente generamos los saldos de apertura en el plan de cuentas
          
          Call GenSaldosAperturaAccessFromSQLServer(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
        End If
      End If
      
     ' Me.MousePointer = vbDefault
   End If
   
End Function


Public Function GenCompAperturaDuplicados(ByVal IdEmpresa As Long, ByVal Rut As String, ByVal Ano As Integer, Optional ByVal GenSaldosAp As Boolean = True) As Boolean
   Dim NumCompAper As Long
   Dim IdCuentaResul As Long
   Dim IdCompAper As Long, IdCompAperTrib As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim Saldo As Double
   Dim Debe As Double
   Dim Haber As Double
   Dim Frm As FrmApertura
   Dim Rc As Integer
   Dim ResDebe As Double
   Dim ResHaber As Double
   Dim SaldoRes As Double
   Dim NAper As Long
   
   GenCompAperturaDuplicados = False
   
   'pedimos al usuario el nº de comp. de apertura, si corresponde, y la cuenta de resultado
   Set Frm = New FrmApertura
   Rc = Frm.FSelectDuplicados(IdEmpresa, Ano, NumCompAper, IdCompAper, IdCuentaResul, IdCompAperTrib)
   Set Frm = Nothing

   If Rc = vbCancel Then   'se arrepintió de generar el comprobante de apertura

      'vemos si no hay comprobante de apertura ya generado
   
      Q1 = "SELECT IdCompAper "
      Q1 = Q1 & " FROM EmpresasAno "
      Q1 = Q1 & " WHERE idEmpresa=" & IdEmpresa & " AND Ano=" & Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         IdCompAper = vFld(Rs("IdCompAper"))
      End If
      
      Call CloseRs(Rs)

      If IdCompAper = 0 Then  'se supone que no hay uno generado (y por lo tanto no hay un comprobante de apertura generado tampoco)
         
         'verifiquemos por si acaso si hay algún comprobante de tipo Apertura.
         'Si es así, se elimina (esto no debiera ocurrir nunca)
         Call VerificaMultiCompApertura
      
         'ahora no hay uno generado, generamos uno en blanco de cada tipo: financiero y tributario, para guardar el número
         Call GenCompAperSinMovs(1, IdEmpresa, Ano, IdCompAperTrib)
         
      End If
   
      Exit Function
   End If
         
   If NumCompAper = 0 Then
      NumCompAper = 1
   End If
   
   'generamos saldos de apertura, tanto financiero como tributario

   If GenSaldosAp Then
      If GenSaldosApertura(IdEmpresa, Rut, Ano, True) = False Then
         Exit Function
      Else
        ' MsgBox1 "Se calcularon los saldos de apertura.", vbInformation
      End If
   End If
   
   '***------ generamos comprobante de apertura financiero  -------***
   
   
   'vemos si hay diferencia en los saldos de apertura financieros
   Q1 = "SELECT Sum(Debe) as SumDebe, Sum(Haber) as SumHaber "
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("SumDebe"))
      TotHaber = vFld(Rs("SumHaber"))
   End If
   
   Call CloseRs(Rs)
   
   Saldo = TotDebe - TotHaber
   
   If Saldo <> 0 Then
   
      'hay diferencia, ajustamos saldo de apertura de la cuenta de resultado
      
      Q1 = "SELECT Debe, Haber FROM Cuentas "
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      ResDebe = 0
      ResHaber = 0
      SaldoRes = 0
      
      If Rs.EOF = False Then
         ResDebe = vFld(Rs("Debe"))
         ResHaber = vFld(Rs("Haber"))
      End If
      
      Call CloseRs(Rs)
      
      SaldoRes = (ResDebe - ResHaber) - Saldo
      
      Q1 = "UPDATE Cuentas SET "
      
      If SaldoRes > 0 Then
         Q1 = Q1 & "  Debe = " & SaldoRes
         Q1 = Q1 & ", Haber = 0"
      Else
         Q1 = Q1 & "  Debe = 0"
         Q1 = Q1 & ", Haber = " & Abs(SaldoRes)
      End If
      
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      
      Call ExecSQL(DbMain, Q1)
   End If
   
   'con los saldos de apertura iguales, generamos comprobante de apertura financiero
   If IdCompAper > 0 And IdCompAperTrib > 0 Then
      Q1 = "  WHERE IdComp = " & IdCompAper
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "MovComprobante", Q1)
      
      Q1 = " WHERE IdComp = " & IdCompAperTrib
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call DeleteSQL(DbMain, "MovComprobante", Q1)
   
   Else           'nuevo comprobante, lo agregamos
   
      'antes verifiquemos por si acaso si hay algún comprobante de tipo Apertura (financiero o tributario).
      'Si es así, se elimina (esto no debiera ocurrir nunca)
      Call VerificaMultiCompApertura

      IdCompAper = GenCompAperSinMovs(NumCompAper, IdEmpresa, Ano, IdCompAperTrib)
      
   End If

   'insertamos movs comprobante, de acuerdo a los saldos de apertura
   Q1 = "INSERT INTO MovComprobante (IdComp, Orden, IdCuenta, Debe, Haber, Glosa, IdEmpresa, Ano)"
   Q1 = Q1 & "  SELECT " & IdCompAper & " As IdComp "
   Q1 = Q1 & ", 1 as Orden, IdCuenta As IdCuenta "
   Q1 = Q1 & ", Debe as Debe, Haber as Haber "
   Q1 = Q1 & ", 'Apertura' as Glosa, " & gEmpresa.id & " As IdEmpresa," & Ano & " As Ano"
   Q1 = Q1 & "  FROM Cuentas WHERE (Debe-Haber) <> 0 "
   Q1 = Q1 & "  AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & "  ORDER BY Cuentas.Codigo "
   
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos totales comprobante de apertura
   Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante "
   Q1 = Q1 & " WHERE IdComp = " & IdCompAper
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano

   Set Rs = OpenRs(DbMain, Q1)

   TotDebe = 0
   TotHaber = 0
   Saldo = 0
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
      Saldo = TotDebe - TotHaber
   End If

   Call CloseRs(Rs)

   'Resultado(Ajuste) para el comprobante de apertura
   'Actualizamos el TotalDebe y TotalHaber del Comprobante de Apertura
   Q1 = "UPDATE Comprobante SET "
   Q1 = Q1 & " Estado=" & EC_APROBADO     'por si lo hubieran anulado
   Q1 = Q1 & ", TipoAjuste=" & TAJUSTE_FINANCIERO
   Q1 = Q1 & ",TotalDebe = " & TotDebe
   Q1 = Q1 & ",TotalHaber=" & TotHaber
   Q1 = Q1 & " WHERE idComp=" & IdCompAper
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)

   If Saldo <> 0 Then      'esto no debiera ocurrir nunca, ya que ya se hizo el ajuste en la tabla de cuentas
      'MsgBox1 "Error al generar comprobante de apertura. Saldo de Debe y Haber no son iguales.", vbExclamation + vbOKOnly
   End If

   Call AddLogComprobantes(IdCompAper, gUsuario.IdUsuario, O_EDIT, Now, EC_APROBADO, NumCompAper, CLng(DateSerial(Ano, 1, 1)), TC_APERTURA, EC_APROBADO, TAJUSTE_FINANCIERO)

   
   '***------ generamos comprobante de apertura trubutario  -------***
   
   
   'vemos si hay diferencia en los saldos de apertura tributario
   Q1 = "SELECT Sum(DebeTrib) as SumDebeTrib, Sum(HaberTrib) as SumHaberTrib "
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("SumDebeTrib"))
      TotHaber = vFld(Rs("SumHaberTrib"))
   End If
   
   Call CloseRs(Rs)
   
   Saldo = TotDebe - TotHaber
   
   If Saldo <> 0 Then
   
      'hay diferencia, ajustamos saldo de apertura de la cuenta de resultado
      
      Q1 = "SELECT DebeTrib, HaberTrib FROM Cuentas "
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      ResDebe = 0
      ResHaber = 0
      SaldoRes = 0
      
      If Rs.EOF = False Then
         ResDebe = vFld(Rs("DebeTrib"))
         ResHaber = vFld(Rs("HaberTrib"))
      End If
      
      Call CloseRs(Rs)
      
      SaldoRes = (ResDebe - ResHaber) - Saldo
      
      Q1 = "UPDATE Cuentas SET "
      
      If SaldoRes > 0 Then
         Q1 = Q1 & "  DebeTrib = " & SaldoRes
         Q1 = Q1 & ", HaberTrib = 0"
      Else
         Q1 = Q1 & "  DebeTrib = 0"
         Q1 = Q1 & ", HaberTrib = " & Abs(SaldoRes)
      End If
      
      Q1 = Q1 & " WHERE IdCuenta = " & IdCuentaResul     'se utiliza la misma cuenta para el comprobante financiero y el tributario
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      Call ExecSQL(DbMain, Q1)
   End If
   

   'insertamos movs comprobante, de acuerdo a los saldos de apertura
   Q1 = "INSERT INTO MovComprobante (IdComp, Orden, IdCuenta, Debe, Haber, Glosa, IdEmpresa, Ano)"
   Q1 = Q1 & "  SELECT " & IdCompAperTrib & " As IdComp "
   Q1 = Q1 & ", 1 as Orden, IdCuenta As IdCuenta "
   Q1 = Q1 & ", DebeTrib as Debe, HaberTrib as Haber "
   Q1 = Q1 & ", 'Apertura' as Glosa, " & gEmpresa.id & " As IdEmpresa," & Ano & " As Ano"
   Q1 = Q1 & "  FROM Cuentas WHERE (DebeTrib-HaberTrib) <> 0 "
   Q1 = Q1 & "  AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Q1 = Q1 & "  ORDER BY Cuentas.Codigo "
   
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos totales comprobante de apertura
   Q1 = "SELECT Sum(Debe) as TotDebe, Sum(Haber) as TotHaber FROM MovComprobante "
   Q1 = Q1 & " WHERE IdComp = " & IdCompAperTrib
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano

   Set Rs = OpenRs(DbMain, Q1)

   TotDebe = 0
   TotHaber = 0
   Saldo = 0
   
   If Rs.EOF = False Then
      TotDebe = vFld(Rs("TotDebe"))
      TotHaber = vFld(Rs("TotHaber"))
      Saldo = TotDebe - TotHaber
   End If

   Call CloseRs(Rs)

   'Resultado(Ajuste) para el comprobante de apertura
   'Actualizamos el TotalDebe y TotalHaber del Comprobante de Apertura
   Q1 = "UPDATE Comprobante SET "
   Q1 = Q1 & " Estado=" & EC_APROBADO     'por si lo hubieran anulado
   Q1 = Q1 & ",TipoAjuste=" & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & ",TotalDebe = " & TotDebe
   Q1 = Q1 & ",TotalHaber=" & TotHaber
   Q1 = Q1 & " WHERE idComp=" & IdCompAperTrib
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)

   If Saldo <> 0 Then      'esto no debiera ocurrir nunca, ya que ya se hizo el ajuste en la tabla de cuentas
     ' MsgBox1 "Error al generar comprobante de apertura. Saldo de Debe y Haber no son iguales.", vbExclamation + vbOKOnly
   End If

   Call AddLogComprobantes(IdCompAper, gUsuario.IdUsuario, O_EDIT, Now, EC_APROBADO, NumCompAper, CLng(DateSerial(Ano, 1, 1)), TC_APERTURA, EC_APROBADO, TAJUSTE_TRIBUTARIO)
   
   GenCompAperturaDuplicados = True
   
End Function

'Pruebas Tiempo
Public Sub CorrigePagadosAñoAnterioresPrueba(ByVal vMsj As Boolean)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim Rs4 As Recordset
   Dim Rs5 As Recordset
   
   Dim i As Integer
   Dim AñoEliminacion As Long
   
   
  ' Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
   
   On Error Resume Next
   
        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If
        
    ERR.Clear
        
         If gDbType = SQL_ACCESS Then
             PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 'PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & Year(fechaemi) & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 
                 PathDbAnoAnt = Replace(PathDbAnoAnt, "\..", "")
             
             If ExistFile(PathDbAnoAnt) Then
               ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
               Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
               
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,NumDoc,Estado,SaldoDoc,TipoLib,TipoDoc,IdEntidad,FEmision "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & " Where estado = " & ED_PAGADO
               Q1 = Q1 & " And FExported > 0 "
               Q1 = Q1 & " And SaldoDoc = 0 "
               Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
               Q1 = Q1 & " And ano = " & gEmpresa.Ano - 1
               
              
               Set Rs = OpenRs(DbAnoAnt, Q1)
               
               Do While Not Rs.EOF
               
                         
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc "
                                 Q1 = Q1 & " From documento "
                                 Q1 = Q1 & " Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & " And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                 Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
                    
                                 Set Rs2 = OpenRs(DbMain, Q1)
                                 
                                 If Rs2.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                                      Call ExecSQL(DbMain, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                                      Call ExecSQL(DbMain, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs2)
                 Rs.MoveNext

                Loop

                Call CloseRs(Rs)
                Call CloseDb(DbAnoAnt)
            End If
      Else
               
           
               Q1 = ""
               Q1 = Q1 & "SELECT IdDoc,NumDoc,Estado,SaldoDoc,TipoLib,TipoDoc,IdEntidad,FEmision "
               Q1 = Q1 & "From documento "
               Q1 = Q1 & " Where estado = " & ED_PAGADO
               Q1 = Q1 & " And FExported > 0 "
               Q1 = Q1 & " And SaldoDoc = 0 "
               Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
               '3405779
               Q1 = Q1 & " And ano = " & gEmpresa.Ano - 1
               '3405779
               Set Rs = OpenRs(DbAnoAnt, Q1)
               
               Do While Not Rs.EOF
               
                         
                                 Q1 = ""
                                 Q1 = Q1 & "SELECT IdDoc "
                                 Q1 = Q1 & " From documento "
                                 Q1 = Q1 & " Where NumDoc = '" & vFld(Rs("NumDoc")) & "' "
                                 Q1 = Q1 & " And TipoLib = " & vFld(Rs("TipoLib"))
                                 Q1 = Q1 & " And TipoDoc = " & vFld(Rs("TipoDoc"))
                                 Q1 = Q1 & " And IdEntidad = " & vFld(Rs("IdEntidad"))
                                 Q1 = Q1 & " and FEmision = " & vFld(Rs("FEmision"))
                                 Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
                                 '3405779
                                 Q1 = Q1 & " And ano = " & gEmpresa.Ano
                                 '3405779
                                 Set Rs2 = OpenRs(DbMain, Q1)
                                 
                                 If Rs2.EOF = False Then
            
                                      Q3 = ""
                                      Q3 = "delete from movDocumento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                                      Q3 = Q3 & " And ano = " & gEmpresa.Ano
                                      Call ExecSQL(DbMain, Q3)
                                     
                                      Q3 = ""
                                      Q3 = "delete from documento "
                                      Q3 = Q3 & " where iddoc = " & vFld(Rs2("iddoc"))
                                      Q3 = Q3 & " And ano = " & gEmpresa.Ano
                                      Call ExecSQL(DbMain, Q3)
                                  End If
                                  
                                  Call CloseRs(Rs2)
                 Rs.MoveNext

                Loop

                Call CloseRs(Rs)
                'Call CloseDb(DbAnoAnt)
                
               
   
    End If
   If vMsj Then
  MsgBox1 "Documentos pagados años anteriores eliminados correctamente.", vbInformation + vbOKOnly
   End If
End Sub
''SF 14202137


#If DATACON = 1 Then
'proceso solo para access
Public Sub CorrigeDocEliminados()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim Rs4 As Recordset
   Dim Rs5 As Recordset
   
   Dim i As Integer
   
  ' Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
   
   On Error Resume Next
   
        #If DATACON = 1 Then
        Dim DbAnoRes As Database
        #End If
        
    ERR.Clear
    
       PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & gEmpresa.Rut & "_R.mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 'PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & Year(fechaemi) & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
                 
                 PathDbAnoAnt = Replace(PathDbAnoAnt, "\..", "")
             
             If ExistFile(PathDbAnoAnt) Then
               ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
               'Set DbAnoRes = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
               
                Call LinkMdbTable(DbMain, PathDbAnoAnt, "Documento", "DocumentoRespaldo", , , ConnStr)
                Call LinkMdbTable(DbMain, PathDbAnoAnt, "MovDocumento", "MovDocumentoRespaldo", , , ConnStr)
                                                  
             
                     Q1 = ""
                     Q1 = Q1 & " SELECT distinct movcomprobante.iddoc "
                     Q1 = Q1 & " from comprobante,movcomprobante"
                     Q1 = Q1 & " Where "
                     
                     'Q1 = Q1 & "Fecha Between " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " and " & CLng(DateSerial(gEmpresa.Ano, 1, 31))
                     'Q1 = Q1 & " and comprobante.idcomp = movcomprobante.idcomp"
                     Q1 = Q1 & " comprobante.idcomp = movcomprobante.idcomp"
                     Q1 = Q1 & " and movcomprobante.iddoc not in (select iddoc from documento)"
                     Q1 = Q1 & " and iddoc > 0"
                     Q1 = Q1 & " And comprobante.IdEmpresa = " & gEmpresa.id
        
                     Set Rs = OpenRs(DbMain, Q1)
                     
                      Do While Not Rs.EOF
                                         
                                Q1 = ""
                                Q1 = Q1 & "SELECT * "
                               ' Q1 = Q1 & "From documento "
                                Q1 = Q1 & "From DocumentoRespaldo "
                                Q1 = Q1 & " Where iddoc = " & vFld(Rs("iddoc"))
                                Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
            
            
                                'Set Rs2 = OpenRs(DbAnoRes, Q1)
                                Set Rs2 = OpenRs(DbMain, Q1)
                                
                                 If Rs2.EOF = False Then
                                 
'                                       Q1 = ""
'                                        Q1 = Q1 & "SELECT DeCentraliz "
'                                       ' Q1 = Q1 & "From documento "
'                                        Q1 = Q1 & "From movcomprobante "
'                                        Q1 = Q1 & " Where iddoc = " & vFld(Rs("iddoc"))
'                                        Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id
'
'
'                                        'Set Rs2 = OpenRs(DbAnoRes, Q1)
'                                        Set Rs3 = OpenRs(DbMain, Q1)
'
'                                        Dim vIdCompCent As Double
'                                        Dim vIdCompPago As Double
'
'                                         If Rs3.EOF = False Then
'
'                                         vIdCompCent = vFld(Rs("DeCentraliz"))
'
'                                         End If
                                    
                                 
                                 
                                 
                                    Q1 = " INSERT INTO DOCUMENTO  Select IdDoc,IdEmpresa,Ano,IdCompCent,IdCompPago,TipoLib,TipoDoc,NumDoc,NumDocHasta,IdEntidad,TipoEntidad "
                                    Q1 = Q1 & " ,RutEntidad,NombreEntidad,FEmision,FVenc,Descrip,Estado,Exento,IdCuentaExento,Afecto,IdCuentaAfecto"
                                    Q1 = Q1 & " ,IVA,IdCuentaIVA,OtroImp,IdCuentaOtroImp,Total,IdCuentaTotal,IdUsuario,FechaCreacion,FEmisionOri"
                                    Q1 = Q1 & " ,CorrInterno,SaldoDoc,FExported,OldIdDoc,DTE,PorcentRetencion,TipoRetencion,MovEdited,OtrosVal"
                                    Q1 = Q1 & " ,FImporF29,NumDocRef,IdCtaBanco,TipoRelEnt,IdSucursal,TotPagadoAnoAnt,FImportSuc,Giro,FacCompraRetParcial"
                                    Q1 = Q1 & " ,IVAIrrecuperable,DocOtrosEnAnalitico,OldIdDocTmp,NumFiscImpr,NumInformeZ,CantBoletas,VentasAcumInfZ,IdDocAsoc"
                                    Q1 = Q1 & " ,PropIVA,ValIVAIrrec,IVAInmueble,FImpFacturacion,CodSIIDTEIVAIrrec,TipoDocAsoc,IVAActFijo,EntRelacionada,NumCuotas"
                                    Q1 = Q1 & " ,CompraBienRaiz,NumDocAsoc,DTEDocAsoc,IdANegCCosto,UrlDTE,CodCtaAfectoOld,CodCtaExentoOld,CodCtaTotalOld,DocOtroEsCargo"
                                    Q1 = Q1 & " ,ValRet3Porc,IdCuentaRet3Porc,Tratamiento "
                                    Q1 = Q1 & " From DocumentoRespaldo "
                                    Q1 = Q1 & " Where IdDoc = " & vFld(Rs("iddoc"))
   
                                    Call ExecSQL(DbMain, Q1)
                                    
                                    Q1 = " INSERT INTO MovDocumento  Select IdMovDoc,IdEmpresa,Ano,IdDoc,IdCompCent,IdCompPago,Orden,IdCuenta,Debe,Haber "
                                    Q1 = Q1 & " ,Glosa,IdTipoValLib,EsTotalDoc,IdCCosto,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld"
                                    Q1 = Q1 & " From MovDocumentoRespaldo "
                                    Q1 = Q1 & " Where IdDoc = " & vFld(Rs("iddoc"))
                                    
                                    Call ExecSQL(DbMain, Q1)
                                    
                                 End If
                                 Call CloseRs(Rs2)
                                        
                                    
                     Rs.MoveNext

                Loop

             Call CloseRs(Rs)
             
               Q1 = "DROP TABLE DocumentoRespaldo"
               Call ExecSQL(DbMain, Q1)
            
               Q1 = "DROP TABLE MovDocumentoRespaldo"
               Call ExecSQL(DbMain, Q1)
                              
    End If
      MsgBox1 "Proceso terminado.", vbInformation + vbOKOnly

End Sub
''SF 14202137

#End If


Public Function CorrigeDocPendientesAñoAnterior(ByVal vMsj As Boolean) As Boolean
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
  
   On Error Resume Next

    ERR.Clear
    
    CorrigeDocPendientesAñoAnterior = False
    
Q1 = ""
Q1 = Q1 & " SELECT iddoc,NUMDOC, TIPODOC, IDENTIDAD, IDEMPRESA FROM DOCUMENTO "
Q1 = Q1 & " WHERE YEAR(FEmision) < " & gEmpresa.Ano
Q1 = Q1 & " And IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
Q1 = Q1 & " AND estado = " & ED_PENDIENTE

Set Rs = OpenRs(DbMain, Q1)

Do While Not Rs.EOF

   Q2 = "UPDATE Documento  "
   Q2 = Q2 & " SET estado = " & ED_CENTRALIZADO
   Q2 = Q2 & " WHERE iddoc = " & vFld(Rs("iddoc"))
   Q2 = Q2 & " AND numdoc = '" & vFld(Rs("numdoc")) & "'"
   Q2 = Q2 & " AND TIPODOC = " & vFld(Rs("TIPODOC"))
   Q2 = Q2 & " AND YEAR(FEmision) < " & gEmpresa.Ano
   Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q2)
   
   
CorrigeDocPendientesAñoAnterior = True
   
   Rs.MoveNext
   
   
Loop
   If vMsj Then
  'MsgBox1 "Documentos duplicados eliminados correctamente.", vbInformation + vbOKOnly
   End If
End Function
'3340329
Public Function GetRemIVAUTM_New(ByVal Mes As Integer, ByVal Ano As Integer, RemUTMMes As Double, RemMesAnte As Double) As Integer
   Dim TotIVACred As Double
   Dim TotIVADeb As Double
   Dim Remanente As Double
   Dim ValUTM As Double
   Dim Fecha As Long
   Dim RemUTM As Double
   Dim RemUTMAnoAnt As Double
   Dim TotRemUTM As Double
   Dim i As Integer, Rc As Integer
   Dim IVAIrrec As Double
   Dim TotIEPDGen As Double, TotIEPDTransp As Double
   Dim IVARetParcial As Double, IVARetTotal As Double
   Dim AjusteIVAMensual As Double
   
   RemUTMMes = 0
   
   GetRemIVAUTM_New = 0
   Rc = GetRemIVAAnoAnt(RemUTMAnoAnt)
   If Rc < 0 Then
      GetRemIVAUTM_New = Rc
      Exit Function
   End If
      
   'OJO redondeamos a dos decimales
   RemUTMAnoAnt = Format(RemUTMAnoAnt, DBLFMT2)
   
   TotRemUTM = RemUTMAnoAnt
   
   For i = 1 To Mes - 1
  
      Fecha = DateSerial(Ano, i + 1, 1)            'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011
      
'      Call GetResIVA(i, Ano, TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
'      IVAIrrec = GetIVAIrrec(i, Ano, i = 1) ' 15 feb 2020
'      Call GetIVARet(i, Ano, IVARetParcial, IVARetTotal, False) ' 15 feb 2010
'
'      AjusteIVAMensual = GetAjusteIVAMensual(i)    ' FCA 28/09/2021
'
'      Remanente = TotIVACred + TotIEPDGen + TotIEPDTransp + AjusteIVAMensual - IVAIrrec - (TotIVADeb - IVARetParcial - IVARetTotal)
''
      If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
         
      
         RemUTM = (RemMesAnte / ValUTM)
         
         'OJO redondeamos a dos decimales
         
         '3321695
         RemUTM = Format(RemUTM, DBLFMT2)
          'RemUTM = (Fix(RemUTM * 100)) / 100
         '3321695
         TotRemUTM = RemUTM
         
         If TotRemUTM < 0 Then
            TotRemUTM = 0
         End If
         
      Else
         MsgBox1 "No se encontró el valor de la UTM para el mes de " & gNomMes(month(Fecha)) & " " & Ano, vbExclamation
         GetRemIVAUTM_New = ERR_VALUTM
         Exit For
      End If
         
      
   Next i
   
   RemUTMMes = TotRemUTM
   
End Function
'3340329


'3410269
Private Sub ModConBoletaVPEE()
Dim Q1 As String
 
   'Modificamos el campo conBoleta del tipo VPEE
   Q1 = "UPDATE TipoDocs "
   Q1 = Q1 & " Set DocBoletas = 1 "
   Q1 = Q1 & " WHERE Diminutivo = 'VPEE'"
   Q1 = Q1 & " AND TipoLib = 2"
   Q1 = Q1 & " AND TipoDoc = 20"
   Call ExecSQL(DbMain, Q1)
 
 
End Sub
' Fin '3410269


Private Sub CampoCodSIIDTE()
Dim Q1 As String
 
   'Modificamos el campo CodSIIDTE ya que debe ser 29 para el impuesto especifico diesel transportista
   Q1 = "UPDATE TipoValor "
   Q1 = Q1 & " Set CodSIIDTE = '29'"
   Q1 = Q1 & " WHERE idTValor = 20"
   Q1 = Q1 & " AND CodSIIDTE = '271'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = ""
   Q1 = Q1 & " IF EXISTS"
   Q1 = Q1 & " (SELECT  1 FROM    sys.columns c"
   Q1 = Q1 & " INNER JOIN sys.types t"
   Q1 = Q1 & " ON t.system_type_id = c.system_type_id"
   Q1 = Q1 & " AND t.user_type_id = c.user_type_id"
   Q1 = Q1 & " WHERE c.name = 'CodSIIDTE'"
   Q1 = Q1 & " AND c.[object_id] = OBJECT_ID(N'movDocumento', 'U')"
   Q1 = Q1 & " AND t.name = 'Char' AND c.max_length = 2 AND c.is_computed = 0)"
   Q1 = Q1 & " BEGIN"
   Q1 = Q1 & " ALTER TABLE movDocumento ALTER COLUMN  CodSIIDTE Char(5);"
   Q1 = Q1 & " END"
   
   Call ExecSQL(DbMain, Q1)
   
End Sub


'633824
Private Sub CampoDocOtroEsCargoRem()
Dim Q1 As String

   'Modificamos el campo DocOtroEsCargo para corregir saldos de rem traspasados del año anterior
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " Set DocOtroEsCargo = 1"
   Q1 = Q1 & " where TipoLib = 4 and TipoDoc = 1"
   Q1 = Q1 & " and idempresa = 2 and year(45291) <= 2023"
   Q1 = Q1 & " and OldIdDoc is not null "
   Q1 = Q1 & " and DocOtroEsCargo = 0"


   Call ExecSQL(DbMain, Q1)


End Sub
'633824
