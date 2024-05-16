Attribute VB_Name = "HyperComun"
Option Explicit

#If DATACON = 2 Then
Public Const DB_MSSQL As Boolean = True    ' Ms SQL Server - 21 ago 2018
#Else
Public Const DB_MSSQL As Boolean = False
#End If

''''''''''''''''''''''
Public gHtmExt As ExtInfo_t

'FRANCA ORIG,FALTABA ESO
'''''''''''''''''''''''''
Public Const P_EMPSEPARADAS = "'EMPSEP'"     ' EMP_SEP
'Public Const API_KEY_TECNOBACK = "RCCYCj9xZl5Wld7jNLEc4Jct4xdaB1v4pLKc1ybd"

#If DATACON = 1 Then
Public DbMain As Database
#Else
Public DbMain As ADODB.Connection
#End If
Public DbMainDate As Double

Public gLexContab As String

Public DbMainAuxx As Database

'canridad de registros para paginamiento en SQL Server
Public Const PAGE_NUMREG = 1000
Public gPageNumReg As Integer


'bases de datos
Public gDbPath As String
Public gDbPathRespaldoTras As String
Public gDbPathAux As String

Public gHRPath As String         ' Path al directorio HR

Public gEmprSeparadas As Boolean  'indica si estamos usando una base de datos con las empresas separadoa (no juntas en uan sola DB)

'máximo crédito art. 33 bis para Activos fijos
Public gMaxUTMCred33 As Double         'hasta 2011: &%= UTM, desde 2012: 500 UTM
Public gMaxUTMCred33_Pesos As Double   'se calcula como gMaxUTMCred33 * UTM
Public gMaxCred33 As Double            'se almacena en tabla ParamEmpresa (MAXCRED33)

' pam: Nueva Instancia: para cuando el programa se llama a si mismo creando una nueva instancia

Public gNuevaInstancia As Boolean

' Items
Public Const PC_NOM = "D3" '  10 -
Public Const PC_COD = "D2" '  75 -
Public Const PC_MAC = "D1" '  30 -
Public Const PC_AUT = "D4" ' 155 - Autorizado
Public Const PC_FEC = "D5" ' 167 - Fecha primera inscripcion
Public Const PC_ULT = "D6" ' 137 - Fecha ultimo uso


Public Const PC_NCOD = "D3"   ' Codigo de licencia
Public Const PC_NIV = "D2"    ' Nivel
Public Const PC_RUT = "D1"    ' RUT de las licencias
Public Const PC_NLIC = "D4"   ' Cantidad de licencias

' Secciones
Public Const PC_EQUIP = "ZCTR1"
Public Const PC_INFO = "LKMR5"

'Orientación de Papel
Public Const ORIENT_VER = 1
Public Const ORIENT_HOR = 2

'Marca de Descontinuado
Public Const MARCA_DESCONTINUADO = "(*)"

'Estado de Entidad
Public Const EE_ACTIVO = 0
Public Const EE_INACTIVO = 1
Public Const EE_BLOQUEADO = 2
Public gEstadoEntidad(EE_BLOQUEADO) As String


'Tipo de Rezones Financieras
Public Const RF_ENDEUDAMIENTO = 1
Public Const RF_LIQUIDEZ = 2
Public Const RF_RENTABILIDAD = 3
Public Const RF_ROTACIONES = 4
Public Const RF_CONSOLIDACION = 5
Public Const RF_OBSOLESCENCIA = 6
Public Const RF_OTROS = 7

'Tipo Ajuste de Comprobante para reportes IFRS
Public Const TAJUSTE_FINANCIERO = 1
Public Const TAJUSTE_TRIBUTARIO = 2
Public Const TAJUSTE_AMBOS = 3

Public Const N_TIPOAJUSTE = TAJUSTE_AMBOS

Public gTipoAjuste(N_TIPOAJUSTE) As String

'Tipo de Contribuyente
Public Const CONTRIB_SAABIERTA = 1           'Soc. Abierta
Public Const CONTRIB_SACERRADA = 2           'Soc. Cerrada
Public Const CONTRIB_SPORACCION = 3          'Soc. por Acción (Encomandita)
Public Const CONTRIB_PRIMCAT = 4             'Soc. Personas primera categoría
Public Const CONTRIB_EMPINDIVIDUALEIRL = 5   'empresario individual (EIRL)
Public Const CONTRIB_EMPINDIVIDUAL = 6       'empresario individual
Public Const CONTRIB_SOCPROFESIONAL = 7      'sociedad de profesionales
Public Const CONTRIB_ESTABPERMANENTE = 8     'establecimiento permanente
Public Const CONTRIB_COMUNIDAD = 9           'comunidad
Public Const CONTRIB_COOPERATIVAS = 10        'cooperativas
Public Const CONTRIB_ORGSINFINESDELUCRO = 11  'org. sin fines de lucro
Public Const MAX_CONTRIB = CONTRIB_ORGSINFINESDELUCRO
'Public Const MAX_CONTRIB = 7

'Franquicias Tributarias
Public Const FRANQ_14BIS = 1                 'Régimen Artículo 14 bis
Public Const FRANQ_LEY18392 = 2              'Ley 18.392 / 19.149
Public Const FRANQ_DL600 = 3                 'D. L. 600
Public Const FRANQ_DL701 = 4                 'D. L. 701
Public Const FRANQ_DS341 = 5                 'D. S. 341
Public Const FRANQ_14TER = 6                 'Régimen Artículo 14 Ter A)
Public Const FRANQ_14QUATER = 7              'Régimen Artículo 14 quater
Public Const FRANQ_RENTAATRIB = 8            'Régimen Renta Atribuida
Public Const FRANQ_SEMIINTEGRADO = 9         'Régimen Semi Integrado
Public Const FRANQ_SOCPROFPRIMCAT = 10       'Soc. Prof. 1ra. Categoría
Public Const FRANQ_SOCPROFSEGCAT = 11        'Soc. Prof. 2da. Categoría
Public Const FRANQ_14ASEMIINTEGRADO = 12     '14 A Régimen Semi Integrado
Public Const FRANQ_PROPYMEGENERAL = 13       '14 D N°3 Régimen Pro Pyme General
Public Const FRANQ_PROPYMETRANSP = 14        '14 D N°8 Régimen Pro Pyme Transp.
Public Const FRANQ_RENTASPRESUNTAS = 15      'Rentas Presuntas
Public Const FRANQ_RENTAEFECTIVA = 16        '14 B N° 1 Renta efectiva sin Balance
Public Const FRANQ_OTRO = 17                 'Otro
Public Const FRANQ_NOSUJETOART14 = 18        'No sujeto art. 14 LIR


Public Const MAX_FRANQ = FRANQ_NOSUJETOART14
Public Const INI_FRANQ2020 = FRANQ_14ASEMIINTEGRADO

'Clasificacion de Entidades
Public Const ENT_CLIENTE = 0
Public Const ENT_PROVEEDOR = 1
Public Const ENT_EMPLEADO = 2
Public Const ENT_SOCIO = 3
Public Const ENT_DISTRIB = 4
Public Const ENT_OTRO = 5
Public Const MAX_ENTCLASIF = ENT_OTRO

Public gClasifEnt(MAX_ENTCLASIF) As String
Public Const SIN_CLASLST = -1
Public Const CON_CLASIF = 1

'Entidad especial para Formulario de Importación
Public Const ENTIMP_RUT = "55.555.555-5"
Public Const ENTIMP_RSOCIAL = "DIN"

'códigos F22 ya no válidos desde 2017
Public Const LSTCODF22_2017 = "101,784,129,645,646,647,648,122,123,844"

'Tipos de depreciaciones de Activos Fijos
Public Const DEP_NORMAL = 1         'depreciación normal
Public Const DEP_ACELERADA = 2      'depreciación acelerada
Public Const DEP_INSTANTANEA = 3    'depreciciación instantánea
Public Const DEP_DECIMAPARTE = 4    'depreciación décima parte
Public Const DEP_DECIMAPARTE2 = 5    'depreciación décima parte 2

Public Const DEP_LEY21210_INST = 1  'Depreciación Ley 21.210 Instantanea e Inmediata
Public Const DEP_LEY21210_ARAUCANIA = 2  'Depreciación Ley 21.210 Araucanía

Public gTipoDepStr(DEP_DECIMAPARTE2) As String
Public gTipoDepLey21210Str(DEP_LEY21210_ARAUCANIA) As String
Public gTipoDepLey21256Str As String


Public lDbRemu As Database

 '3236967
Type FmtImp_Emp
   Campo As String
   Formato As String
End Type
 
Public gImpEmpresas() As FmtImp_Emp

Type Oficina_t
   Nombre   As String
   Rut      As String
End Type
Public gOficina As Oficina_t  ' Info de la oficina de contabilidad


'funciones que se habilitan segun versión
Public gFunciones As Funciones_t

Type Funciones_t
   RazFinancieras As Boolean           'razones financieras
   ActivoFijo As Boolean               'activo fijo
   ExpFUT As Boolean                   'exportación a FUT
   OtrosInformes As Boolean            'Otros informes: cheques anulados
   DetDocReten As Boolean              'Detalle de Documentos de Libro Retenciones
   DetSaldoApertura As Boolean         'Detalle saldo de apertura por Entidad
   ComprobanteResumido As Boolean      'Permite imprimir comprobantes en forma resumida
   ExpImpLibrosAux As Boolean          'Exportación e importación de libros auxiliares (compras, ventas, retenciones) con base de datos
   ExpImpLibrosAuxFile As Boolean      'Exportación e importación de libros auxiliares (compras, ventas, retenciones) con archivo de texto separado por tabulaciones
   ExpPlanCuentas As Boolean           'Exportación del Plan de Cuentas
   ExpHRCertificados As Boolean        'Exportación a HR Certificados
   ExpHRForm22 As Boolean              'Exportación a Form 22
   PrtCheque As Boolean                'Imprimir cheques
   ImportRemu As Boolean               'importar desde FairPay2
   IFRS As Boolean                     'Reportes IFRS
   IFRS_Ejecutivo As Boolean           'Reporte IFRS Ejecutivo
   IFRS_BalanceTributario As Boolean   'Reporte IFRS Balance tributario u 8 columnas
   NuevoTraspasoIVA As Boolean         'traspaso a HR-IVA con una tabla Access que genera este sistema
   NuevoTraspasoForm22 As Boolean      'traspaso a HR-Form22 con una tabla Access que genera este sistema
   ImportRetenciones As Boolean        'importacion de retenciones
   ImportComprobantes As Boolean       'importación de comprobantes
   AuditoriaInterna As Boolean         'auditoria interna
   ControlContrib As Boolean           'control contribuyentes: activo o no activo
   ExpLibCompVentasSII As Boolean      'exportación de libros de compras y ventas en formato SII
   ProporcionalidadIVA As Boolean      'proporcionalidad del IVA
   ActFijoFinanciero As Boolean        'Activo Fijo Financiero IFRS
   RepActFijoFinanciero As Boolean     'Reporte Activo Fijo Financiero IFRS
   LibroCaja As Boolean                'Libro de Caja Ingresos y Egresos
   DocCuotas As Boolean                'cuotas de documentos
   OtrosIngEgresos As Boolean          'importación a libro de caja de otros ingresos y egresos
   AjustesExtraLibCaja  As Boolean     'Ajustes extra Libro de Caja
End Type

Type Entidad_t
   Rut         As String
   NotValidRut As Boolean
   Nombre      As String
   'Descrip    As String
   Codigo      As String
   id          As Long
   Estado      As Integer
   Clasif      As Integer
   email       As String
End Type



' Tipos en ParamEmpresa
Public Const TPE_DBINFO = "'DBINFO'"


'archivo INI
Public gIniFile As String
Public gCfgFile As String

Public gDebug As Integer
Public gChecked As Boolean

Public CodCuentaSelec As Long

' String de conexión a la base Comun
Public gComunConnStr As String

'usuarios y privilegios
Public Const PRV_ADMIN = &H1FFFFF

'nombre de usuario del administrador para cada aplicación
Public gAdmUser As String

'BASES DE DATOS
Public Const OPENMDB_ADM = 1
Public Const OPENMDB_EMP = 2
Public gEmpresa As Empresa_t
''Public gRegEmpAnoAnt As RegEmpresa_t

Public gEmprHR As EmpresaHR_t

'Nombre Base de Datos Vacia
'Public Const BD_COMUN = "LexContab.mdb"
Public Const BD_COMUN = "LPContab.mdb"
Public Const BD_IFRS = "TblIFRS.mdb"
Public Const BD_IFRS_50 = "TblIFRS_50.mdb"

'Nombre Base de Datos Vacia
Public Const BD_VACIA = "EmpresaVacia.mdb"

'CARACTERSTICA MONEDA
Global Const MON_NACION = 0
Global Const MON_VUNICO = 1
Global Const MON_VMES = 2
Global Const MON_VDIA = 3

'tipos de comprobantes
Public Const TC_EGRESO = 1
Public Const TC_INGRESO = 2
Public Const TC_TRASPASO = 3
Public Const TC_APERTURA = 4

Public Const N_TIPOCOMP = 4

Public gTipoComp(N_TIPOCOMP) As String

'estados comprobantes
Public Const EC_ANULADO = 1
Public Const EC_APROBADO = 2
Public Const EC_PENDIENTE = 3
Public Const EC_ERRONEO = 4

Public Const N_ESTADOCOMP = 4

Public Const EC_ELIMINADO = -1    'sólo para el informe de auditoría

Public gEstadoComp(N_ESTADOCOMP) As String

'Estado Documento
Public Const ED_CENTPAG = -1    'sólo para combos de selección de docs con estado Centralizados y Pagados
Public Const ED_ANULADO = 1
Public Const ED_APROBADO = 2
Public Const ED_CENTRALIZADO = 3
Public Const ED_PAGADO = 4
Public Const ED_PENDIENTE = 5

Public Const MAX_ESTADODOC = ED_PENDIENTE

Public gEstadoDoc(MAX_ESTADODOC) As String

'estado Cuentas
Public Const ECTA_INACTIVA = 0
Public Const ECTA_ACTIVA = 1

'clasificación de las cuentas (Activo, Pasivo, Capital, etc.)
Public Const CLASCTA_ACTIVO = 1
Public Const CLASCTA_PASIVO = 2
Public Const CLASCTA_RESULTADO = 3
Public Const CLASCTA_ORDEN = 4
Public Const MAX_CLASCTA = CLASCTA_ORDEN
Public gClasCta(MAX_CLASCTA) As String

'ATRIBUTOS DE CUENTA
Public Const ATRIB_CONCILIACION = 1
Public Const ATRIB_CAPITALPROPIO = 2
Public Const ATRIB_ACTIVOFIJO = 3
Public Const ATRIB_RUT = 4
Public Const ATRIB_CCOSTO = 5
Public Const ATRIB_AREANEG = 6
Public Const ATRIB_CAJA = 7
Public Const ATRIB_14TER = 8
Public Const ATRIB_PERCEPCIONES = 9
Public Const MAX_ATRIB = ATRIB_PERCEPCIONES
Public gAtribCuentas(MAX_ATRIB) As Atrib_t

'Tipos de Capital Propio
Public Const CAPPROPIO_ACTIVO_NORMAL = 1
Public Const CAPPROPIO_ACTIVO_VALINTO = 2
Public Const CAPPROPIO_ACTIVO_COMPACTIVO = 3
Public Const CAPPROPIO_PASIVO_EXIGIBLE = 4
Public Const CAPPROPIO_PASIVO_NOEXIGIBLE = 5

Public Const MAX_CAPPROPIO = CAPPROPIO_PASIVO_NOEXIGIBLE

'libros y documentos
Public Const LIB_COMPRAS = 1
Public Const LIB_VENTAS = 2
Public Const LIB_RETEN = 3
Public Const LIB_REMU = 4
Public Const LIB_OTROS = 5   'para agrupar otros docs
Public Const LIB_CAJAING = 6  'para otros ingresosal libro de caja
Public Const LIB_CAJAEGR = 7  'para otros egresos al libro de caja
Public Const LIB_OTROFULL = 8  'para otros egresos al libro de caja

Public CodTipoLib As Long

'Tipo IVA Retenido
Public Const IVARET_PARCIAL = 1     'IVA Retenido Parcial
Public Const IVARET_TOTAL = 2       'IVA Retenido Total


'Tipo de Valor según Libro
'libro de compras
Public Const LIBCOMPRAS_AFECTO = 1
Public Const LIBCOMPRAS_EXENTO = 2
Public Const LIBCOMPRAS_TOTAL = 3
Public Const LIBCOMPRAS_IVACREDFISC = 4
Public Const LIBCOMPRAS_OTROSIMP = 5
Public Const LIBCOMPRAS_IVARETENIDO = 6               'ELIMINADO!
Public Const LIBCOMPRAS_ANTICIPOS = 7                 'ELIMINADO!
Public Const LIBCOMPRAS_IVAIRREC = 8                  'IVA Irrecuperable
Public Const LIBCOMPRAS_IVAACTFIJO = 9
Public Const LIBCOMPRAS_IVAIRRACTFIJO = 10            'ELIMINADO!  IVA Irrecuperable Activo Fijo
Public Const LIBCOMPRAS_IVARETPARC = 11               'IVA Retenido Parcial (Descontinuado)
Public Const LIBCOMPRAS_IVARETTOT = 12                'IVA Retenido Total
'Public Const LIBCOMPRAS_IVAANTICIPADO = 13           'ELIMINADO!
Public Const LIBCOMPRAS_IMPPISCO = 14                 'Impto. Adic. Art.42: Pisco, Licores, Wisky y Aguar
Public Const LIBCOMPRAS_IMPVINOS = 15                 'Impto. Adic. Art.42: Vinos, Champaña, Chichas (tas
Public Const LIBCOMPRAS_IMPCERVEZA = 16               'Impto. Adic. Art.42: Cervezas (tasa 15%)
Public Const LIBCOMPRAS_IMPBEBANALC = 17              'Impto. Adic. Art.42: Bebidas Analcohólicas (tasa 1
Public Const LIBCOMPRAS_ILANOTASDEB = 18              'ILA por Notas de Débito recibidas SINUSO!
Public Const LIBCOMPRAS_ILANOTASCRED = 19             'ILA por Notas de Crédito recibidas SINUSO!
Public Const LIBCOMPRAS_IVAANTICIPHARINA = 20         'IVA anticipado del periodo Harina
Public Const LIBCOMPRAS_IVAANTICIPCARNE = 21          'IVA anticipado del periodo Carne
Public Const LIBCOMPRAS_IMPESPDIESEL = 22             'Impuesto específico Diesel
Public Const LIBCOMPRAS_IMPESPPETRGRAL = 23           'Impuesto específico Petróleo Diesel General ELIMINADO!
Public Const LIBCOMPRAS_IMPESPDIESELTRANS = 24        'Impuesto específico Diesel Transportista
Public Const LIBCOMPRAS_IMPESPPETRGENCF = 25          'Impuesto Especifico Petróleo Diesel General CF ELIMINADO!
Public Const LIBCOMPRAS_IMPESPPETRCARGACF = 26        'Impuesto Especifico Petróleo Diesel Trans. Carga CF ELIMINADO!
Public Const LIBCOMPRAS_IMPESPPETRGENSINCF = 27       'Impuesto Específico Petróleo Diesel General Sin Derecho a CF ELIMINADO!
Public Const LIBCOMPRAS_IMPESPPETRCARGASINCF = 28     'Impuesto Específico Petróleo Diesel Trans. Carga Sin Derecho a CF ELIMINADO!
Public Const LIBCOMPRAS_ILABEDANALCAZUCAR = 29        'ILA por Bebidas Analcoholicas con elevado cont. Azúcar ELIMINADO!
Public Const LIBCOMPRAS_IVAADQCONSTINMUEBLES = 30     'IVA por Adq. o Const. Inmuebles ELIMINADO!
Public Const LIBCOMPRAS_IVAIRREC1 = 31                'IVA Irrecuperable: 1.Compras destinadas a generar operaciones no gravadas o exentas
Public Const LIBCOMPRAS_IVAIRREC2 = 32                'IVA Irrecuperable: 2.Facturas de proveedores registradas fuera de plazo
Public Const LIBCOMPRAS_IVAIRREC3 = 33                'IVA Irrecuperable: 3.Gastos rechazados
Public Const LIBCOMPRAS_IVAIRREC4 = 34                'IVA Irrecuperable: 4.Entregas gratuitas (premios, bonificaciones etc.) recibidas
Public Const LIBCOMPRAS_IVAIRREC9 = 35                'IVA Irrecuperable: 9.Otros
Public Const LIBCOMPRAS_IVARETPARCTRIGO = 36          'IVA Retenido Parcial Trigo
Public Const LIBCOMPRAS_IVARETPARCMADERA = 37         'IVA Retenido Parcial Madera
Public Const LIBCOMPRAS_IVARETPARCGANADO = 38         'IVA Retenido Parcial Ganado
Public Const LIBCOMPRAS_IVARETPARCLEGUMBRES = 39      'IVA Retenido Parcial Legumbres
Public Const LIBCOMPRAS_IVARETPARCARROZ = 40          'IVA Retenido Parcial Arroz
Public Const LIBCOMPRAS_IVARETPARCSILVESTRES = 41     'IVA Retenido Parcial Silvestres
Public Const LIBCOMPRAS_IVARETPARCHIDROBIO = 42       'IVA Retenido Parcial Hidrobiológicas
Public Const LIBCOMPRAS_IVARETPARCFAMBPASAS = 43      'IVA Retenido Parcial Frambuezas y Pasas
Public Const LIBCOMPRAS_IVARETTOTCHATARRA = 44        'IVA Retenido Total Chatarra
Public Const LIBCOMPRAS_IVARETTOTPPA = 45             'IVA Retenido Total PPA
Public Const LIBCOMPRAS_IVARETTOTCONSTR = 46          'IVA Retenido Total Construcción
Public Const LIBCOMPRAS_IVARETTOTCARTONES = 47        'IVA Retenido Total Cartones
Public Const LIBCOMPRAS_IVAFAENACARNE = 48            'IVA Anticipado Faenamiento Carne
Public Const LIBCOMPRAS_IMPPIEDRASPREC = 49           'Impto. Joyas, Piedras Prec., Pieles Finas
Public Const LIBCOMPRAS_IMPALFOMBRAS = 50             'Imp. Adicional (alfombras, tapices, casas rodantes, caviar)
Public Const LIBCOMPRAS_IVARETORO = 51                'IVA Retenido Oro
Public Const LIBCOMPRAS_IMPPIROTECNIA = 52            'Impuesto Adicional (Pirotecnia)
Public Const LIBCOMPRAS_IVAMARGCOM = 53               'IVA de Margen de Comercialización.
Public Const LIBCOMPRAS_IMPGASOLINA = 54              'Impuesto Específico Gasolina
Public Const LIBCOMPRAS_IVAMARGCOMPREPAGO = 55        'IVA de Margen de Comer. de Inst. de Prepago
Public Const LIBCOMPRAS_IMPGASNATURAL = 56            'Impuesto Gas Natural Comprimido
Public Const LIBCOMPRAS_IMPGASLIQ = 57                'Impuesto Gas Licuado de Petróleo
Public Const LIBCOMPRAS_IMPSUPLEMENTEROS = 58         'Imp. Retenido Suplementeros Art. 74 n° 5, LIR
Public Const LIBCOMPRAS_IMPBEDANALCAZUCAR = 59        'Imp. Bebidas analcohólicas con alto contenido de azúcar


Public Const LIBCOMPRAS_NUMOTROSIMP = LIBCOMPRAS_IMPBEDANALCAZUCAR - LIBCOMPRAS_IVAIRREC + 1

'libro de ventas
Public Const LIBVENTAS_AFECTO = 1
Public Const LIBVENTAS_EXENTO = 2
Public Const LIBVENTAS_TOTAL = 3
Public Const LIBVENTAS_IVADEBFISC = 4
Public Const LIBVENTAS_OTROSIMP = 5
Public Const LIBVENTAS_IVARETENIDO = 6                         'ELIMINADO!
Public Const LIBVENTAS_RETENCIONES = 7                         'ELIMINADO!
Public Const LIBVENTAS_REBAJA65 = 8
Public Const LIBVENTAS_IVAIRREC = 9                            'ELIMINADO!
Public Const LIBVENTAS_IVARETPARC = 10                         'IVA Retenido Parcial
Public Const LIBVENTAS_IVARETTOT = 11                          'IVA Retenido Total
Public Const LIBVENTAS_RETMARGENCOM = 12                       'Retención márgen de comercialización
'Public Const LIBVENTAS_RETANTCAMBIOSUJ = 13                   'ELIMINADO!
Public Const LIBVENTAS_IMPPISCO = 14                           'Impto. Adic. Art.42: Pisco, Licores, Whisky y Aguard.
Public Const LIBVENTAS_IMPVINOS = 15                           'Impto. Adic. Art.42: Vinos, Champaña, Chichas (tas
Public Const LIBVENTAS_IMPCERVEZA = 16                         'Impto. Adic. Art.42: Cervezas (tasa 15%)
Public Const LIBVENTAS_IMPBEBANHALC = 17                       'Impto. Adic. Art.42: Bebidas Analcohólicas (tasa 1
Public Const LIBVENTAS_ILANOTASDEB = 18                        'ILA por Notas de Débito emitidas
Public Const LIBVENTAS_ILANOTASCRED = 19                       'ILA por Notas de Crédito emitidas
Public Const LIBVENTAS_IMPART37E = 20                          'Impto. Adicional Art.37  e) h) i) l)
Public Const LIBVENTAS_IMPART37J = 21                          'Impto. Adicional Art.37  j)
'Public Const LIBVENTAS_RETANTCAMBIOSUJHARINA = 0              'Antes era 22 pero se utilizó por ERROR para el anticipo de la Carne. Entonces se asigó el 22 al IVA Anticipado Carne. No existe et. Anticipo Harina
Public Const LIBVENTAS_IVAANTICIPADOCARNE = 22                 'IVA Anticipado Carne
Public Const LIBVENTAS_RETANTCAMBIOSUJCARNE = 23               'Retención anticipo cambio sujeto
Public Const LIBVENTAS_ILABEDANALCAZUCAR = 24                  'ILA por Bebidas Analcoholicas con elevado cont. Azúcar
Public Const LIBVENTAS_IVAADQCONSTINMUEBLES = 25               'IVA por Adq. o Const. Inmuebles
Public Const LIBVENTAS_JOYAS = 26                              'Imp. Adicional Joyas, piedras preciosas, pieles finas
Public Const LIBVENTAS_IVAANTICIPADOHARINA = 27                'IVA Anticipado Harina, se agrega 31/01/17
Public Const LIBVENTAS_IVA_ANTICIP_FAENACARNE = 28             'IVA Anticipado Faenamiento Carne
Public Const LIBVENTAS_IVA_RETPARCIAL_LEGUMBRES = 29           'IVA Retenido Parcial Legumbres
Public Const LIBVENTAS_IVA_RETTOTAL_LEGUMBRES = 30             'IVA Retenido Total Legumbres
Public Const LIBVENTAS_IVA_RETTOTAL_SILVESTRES = 31            'IVA Retenido Total Silvestres
Public Const LIBVENTAS_IVA_RETPARCIAL_GANADO = 32              'IVA Retenido Parcial Ganado
Public Const LIBVENTAS_IVA_RETTOTAL_GANADO = 33                'IVA Retenido Total Ganado
Public Const LIBVENTAS_IVA_RETPARCIAL_MADERA = 34              'IVA Retenido Parcial Madera
Public Const LIBVENTAS_IVA_RETTOTAL_MADERA = 35                'IVA Retenido Total Madera
Public Const LIBVENTAS_IVA_RETPARCIAL_TRIGO = 36               'IVA Retenido Parcial Trigo
Public Const LIBVENTAS_IVA_RETTOTAL_TRIGO = 37                 'IVA Retenido Total Trigo
Public Const LIBVENTAS_IVA_RETPARCIAL_ARROZ = 38               'IVA Retenido Parcial Arroz
Public Const LIBVENTAS_IVA_RETTOTAL_ARROZ = 39                 'IVA Retenido Total Arroz
Public Const LIBVENTAS_IVA_RETPARCIAL_HIDROBIOLOGICAS = 40     'IVA Retenido Parcial Hidrobiológicas
Public Const LIBVENTAS_IVA_RETTOTAL_HIDROBIOLÓGICAS = 41       'IVA Retenido Total Hidrobiológicas
Public Const LIBVENTAS_IVA_RETTOTAL_CHATARRA = 42              'IVA Retenido T
Public Const LIBVENTAS_IVA_RETTOTAL_PPA = 43                   'IVA Retenido Total PPA
Public Const LIBVENTAS_IVA_RETTOTAL_CARTONES = 44              'IVA Retenido Total Cartones
Public Const LIBVENTAS_IVA_RETPARCIAL_BERRIES = 45             'IVA Retenido Parcial Berries
Public Const LIBVENTAS_IVA_RETTOTAL_BERRIES = 46               'IVA Retenido Total Berries
Public Const LIBVENTAS_FACT_COMPRA_SIN_RET = 47             'Factura de compra sin Retención
Public Const LIBVENTAS_IVA_RET_FACT_INICIO = 48                'IVA Retenido Factura de Inicio
Public Const LIBVENTAS_IMPDIESEL = 49                          'Impuesto Específico Diesel
Public Const LIBVENTAS_IMPGASOLINA = 50                        'Impuesto Específico Gasolina



'ERROR

Public Const LIBVENTAS_NUMOTROSIMP = LIBVENTAS_IVAADQCONSTINMUEBLES - LIBVENTAS_REBAJA65 + 1
Public Const LIBVENTAS_INIIMPADIC = LIBVENTAS_IMPPISCO

Public Const MAX_COLOTROIMP = LIBCOMPRAS_NUMOTROSIMP


'Libro de Caja
Public Const TOPERCAJA_INGRESO = 1
Public Const TOPERCAJA_EGRESO = 2

Public gTipoOperCaja(TOPERCAJA_EGRESO) As String

Public Const BASELIBCAJA_RETEN = 100                 'para diferenciar documentos de egresos: Compras, Retenciones y otros egresos . Se suma BASELIBCAJA_RETEN al idTipoDoc cuando es de retenciones
Public Const BASELIBCAJA_INGEGR = 200                'para diferenciar documentos de ingreso: Ventas y otros ingresos. Se suma BASELIBCAJA_INGEGR al idTipoDoc cuando es de otros ingresos o egresos

Public Const TOPERCAJA_LOCK = 100                     'para LockAction de Libro de caja

'Tipos Para otros ingresos y egresos al libro de caja
Public Const LIBCAJA_OTROSING = 1                     'Para otros ingresos al libro de caja
Public Const LIBCAJA_OTROSEGR = 2                     'Para otros egresos al libro de caja

Public gTipoDocCajaOtros(LIBCAJA_OTROSEGR) As String

'Diminutivos para otros ingresos y egresos Libro de Caja
Public Const TDOC_OTROSINGRESOS = "OIN"
Public Const TDOC_OTROSEGRESOS = "OEG"

'Tipos Para otros ingresos y egresos Saldo Inicial al libro de caja
Public Const LIBCAJA_OTROSINGINI = 3                     'Para otros ingresos saldo inicial al libro de caja
Public Const LIBCAJA_OTROSEGRINI = 4                     'Para otros egresos saldo inicial al libro de caja

'Diminutivos para Saldo Inicial ingresos y egresos Libro de Caja
Public Const TDOC_OTROSINGRINI = "OII"     'Saldo Inicial Ingresos
Public Const TDOC_OTROSEGRINI = "OEI"       'Saldo Inicial Egresos


'ocultar imp adicionales

Public gOcultarImpAdicDescont As Integer    'opción de ocultar impuestos adicionales en el detalle de un documento

'Tipos de IVA Irrecuperable Libro de Compras
Type IvaIrrec_t
   CodImpSII As Integer
   Descrip As String
End Type

Public Const MAX_TIPOIVAIRREC = 5

Public gTipoIvaIrrec(MAX_TIPOIVAIRREC) As IvaIrrec_t


'libro de retenciones
Public Const LIBRETEN_HONORSINRET = 1
Public Const LIBRETEN_BRUTO = 2
Public Const LIBRETEN_IMPUESTO = 3
Public Const LIBRETEN_NETO = 4
Public Const LIBRETEN_RET3PORC = 5

'libro remuneraciones
Public Const LIBREMU_VALOR = 1
Public Const LIBREMU_TOTAL = 2

'libro Otros documentos
Public Const LIBOTROS_VALOR = 1
Public Const LIBOTROS_TOTAL = 2


Public Const MAX_TIPOVALLIB = 50

'Lock Action
Public Const LK_COMPROBANTE = 1000     'para lock de comprobante
Public Const LK_COMPTIPO = 1001        'para lock de comprobante tipo

Public Const LK_EXPLIBROS = 1002       'para lock de exportación de libros
Public Const LK_IMPLIBROS = 1003       'para lock de importación de libros
Public Const LK_EXPENTIDADES = 1004    'para lock de Exportación de entidades
Public Const LK_IMPENTIDADES = 1005    'para lock de importación de entidades


'diminutivo tipodocs para valiaciones libro de ventas
Public Const TDOC_FACEXENTA = "FCE"
Public Const TDOC_FAVEXENTA = "FVE"
Public Const TDOC_BOLVENTA = "BOV"
Public Const TDOC_BOLEXENTA = "BOE"
Public Const TDOC_BOLVENTAEX = "BEX"
Public Const TDOC_DEVVENTABOL = "DVB"
Public Const TDOC_MAQREGISTRADORA = "MRG"
Public Const TDOC_VALEPAGOELECTR = "VPE"
Public Const TDOC_VENTASINDOC = "VSD"

'2814014 pipe
Public Const TDOC_VALVENTAEX = "VPEE"
'fin 2814014

' Datos del tipo OFICINA en Param
Public Const TOF_NOMBRE = 1
Public Const TOF_RUT = 2

' Tablas que se linkean
Public Const T_CODACTIV = 0   ' CodActiv
Public Const T_EMPRESAS = 1   ' Empresas
Public Const T_EMPANO = 2     ' EmpresasAno
Public Const T_EQUIV = 3      ' Equivalencia
Public Const T_IMPTO = 4      ' Impuestos
Public Const T_MONEDA = 5     ' Monedas
Public Const T_PARAM = 6      ' Param
Public Const T_PAVAN = 7      ' PlanAvanzado
Public Const T_PBAS = 8       ' PlanBásico
Public Const T_PINT = 9       ' PlanIntermedio
Public Const T_REG = 10       ' Regiones
Public Const T_TIMB = 11      ' PlanBásico
Public Const T_TVALOR = 12    ' TipoValor
Public Const T_USER = 13      ' Usuarios
Public gLnkTabla(T_USER) As String


'operaciones sobre un objeto
Public Const O_NEW = 1
Public Const O_EDIT = 2
Public Const O_VIEW = 3
Public Const O_PRINT = 4
Public Const O_SELECT = 5
Public Const O_SELEDIT = 6
Public Const O_SELCOPY = 7

Public Const O_IMPORT = 10
Public Const O_EXPORT = 11

Public Const O_DELETE = 20

Public Const NEGNUMFMT = "#,##0;(#,##0)"
Public Const NEGBL_NUMFMT = "#,###;(#,###)"

'font
Public gUseCourier As Integer

'Para usuarios
Public gUsuario As Usuario_t

'Plan de cuentas actual de la empresa
Public gPlanCuentas As String


'Índices
Public gIVA As Double               '0,19
Public gImpPrimCategoria As Double  '0,25
Public gCredArt33 As Single         '0,04   'crédito activos fijos (desde 1 ene 2012 hasta 30 sept 2014)
Public gCredArt33_2014 As Single    '0,08   'crédito activos fijos (desde 1 oct 2014 hasta 30 sept 2015)
Public gCredArt33_2015 As Single    '0,06   'crédito activos fijos (desde 1 oct 2015)

'3387590 FFV
Public gAfectoCero As Integer   '0 = No Permite valor 0 / 1 = Permite valor 0 en afecto libro compra y ventas Factura Afectas

'impto retención
Public Const IMPRET_NAC = 1         '0,1     (10%)
Public Const IMPRET_EXT = 2         '0,35    (35%)
Public Const IMPRET_OTRO = 3        'se calcula a partir del bruto y el impuesto
Public gImpRet(IMPRET_OTRO) As Single

'Indices financieros
Type Indices_t
   PuntosIPC As Double
   VarIpc As Double
   FactorCM As Double
End Type

Public gIndices(12) As Indices_t
Public gFactorActAnual(12, 12) As SII_Fact_t

 Public FProPymeGeneral As Boolean      '14 D N°3 Régimen Pro Pyme General
 Public FProPymeTransp As Boolean       '14 D N°8 Régimen Pro Pyme Transp

'empresa
Type Empresa_t
   id          As Long
   Rut         As String
   NombreCorto As String
   RazonSocial As String   ' o ApPaterno
   ApMaterno   As String
   Nombre      As String
   Telefono    As String
   Fax         As String
   Region      As Integer
   Comuna      As String
   Ciudad      As String
   Direccion   As String
   Giro        As String
   CodActEcono As String
   RepConjunta As Boolean
   RutRepLegal1 As String
   RepLegal1   As String
   RutRepLegal2 As String
   RepLegal2   As String
   Ano         As Integer
   CopyAno     As Integer
   FCierre     As Long
   FApertura   As Long
   Opciones    As Long
   ConnStr     As String
   TieneAnoAnt As Boolean
   TieneAnoAntAccess As Boolean  'Si es SQL Server, indica que tiene año aterior en Access
   DebeGenCompAp As Boolean   'indica que debe ofrecer generar comprobante de apertura dado que es la primera vez que ingresa al año, tiene año anterior, y no ha generado comprobante de apertura. Si no, es falso
   RutDisp     As String      'Rut que se usa para desplieque en membrete reportes (se hizo exclusivamente para la Asociación De AFP que tienen varias empresas con mismo RUT)
   TipoContrib As Integer
   Franq14Ter  As Boolean     'Franquicia 14 Ter
   RentaAtribuida As Boolean  'Regimen Renta Atribuida
   SemiIntegrado  As Boolean  ' Regimen Semi Integrado
   SocProfSegCat As Boolean   'Sociedad de Profesionales segunda categoría
   R14ASemiIntegrado As Boolean  '14 A Régimen Semi Integrado
   ProPymeGeneral As Boolean     '14 D N°3 Régimen Pro Pyme General
   ProPymeTransp As Boolean      '14 D N°8 Régimen Pro Pyme Transp
   RentasPresuntas As Boolean    'Rentas Presuntas
   RentaEfectiva As Boolean      '14 B N° 1 Renta efectiva sin Balance
   RegimenOtro As Boolean        'Otro
   NoSujetoArt14 As Boolean      'No sujeto art. 14 LIR
   ObligaLibComprasVentas As Boolean
   email       As String
   ObsDTE      As String
   RutFirma      As String
   '2861570
   RutContador As String '2861570 rut de contador
   'fin 2861570
   '2929879
   Villa      As String
   CodArea    As String
   Celular    As String
   'fin 2929879
   'SF 14691055
   NuevoAno As Boolean
   'SF 14691055
End Type

Type EmpresaHR_t
   EmpConta As Empresa_t
   ApMaterno   As String
   Region      As String
   NroCalle    As String
   NroDepto    As String
   NombContador As String
   RutContador As String
   DirPostal   As String
   ComunaPostal As String
   TipoContrib As Integer
   TransaBolsa As Boolean
   Franquicias(MAX_FRANQ) As Boolean
End Type




'ESTRUCTURAS
Type Usuario_t
   Rc          As Integer
   IdUsuario   As Integer
   Nombre      As String
   Priv        As Long
   idPerfil    As Integer
   ClaveACtual As Long
   NombreLargo As String
   
End Type

Type Monedas_t
   id          As Integer
   Descrip     As String
   Simbolo     As String
   DecInf      As Single
   DecVenta    As Single
   Caract      As Byte
   FormatInf   As String
   FormatVenta As String
   EsFijo      As Boolean
End Type

Public gMonedas() As Monedas_t

Type Atrib_t
   Nombre         As String
   NombreCorto    As String
End Type

'Tipos de Valores y tipos de documentos
Type TipoValLib_t
   id             As Long        'id en la tabla TipoValor (autonumber)
   TipoLib        As Integer
   TipoDoc        As String
   TipoValLib     As Integer     'Campo código (id de tipo val por libro)
   Nombre         As String      'Campo Valor
   TitCompleto    As String
   Diminutivo     As String
   Atributo       As String
   Multiple       As Boolean     'indica si es posible tener múltiples movimientos de este tipo de valor en un mismo documento
   CodF29         As Integer
   orden          As Integer
   Tasa           As Single
   TasaFija       As Boolean
   EsRecuperable  As Boolean
   CodSIIDTE      As String
   Descontinuado  As Boolean
   TipoIVARetenido    As Integer     'IVARET_PARCIAl o IVARET_TOTAL
End Type

Public gTipoValLib() As TipoValLib_t      'tipo de movimiento en documentos (Neto, IVA, Bruto,etc.) para desglose de valores en libros de compras y ventas (Tabla param, Tipo "TIPOMOVDOC")

Type NTipoVal_t                  'estructura para contar cuantas veces aparece un TipoValLib en un documento
   IdTipoValLib   As Long
   Count          As Integer
   Valor          As Double
   NombreCampo    As String
End Type

Public Const VAL_NOPERMITIDO = 0
Public Const VAL_OPCIONAL = 1
Public Const VAL_OBLIGATORIO = 2

Type TipoDoc_t
   id                As Long
   TipoLib           As Integer
   TipoDoc           As Integer
   Nombre            As String
   Diminutivo        As String
   Atributo          As String
   TieneAfecto       As Integer        'VAL_NOPERMITIDO, VAL_OPCIONAL y VAL_OBLIGATORIO
   TieneExento       As Integer        'VAL_NOPERMITIDO, VAL_OPCIONAL y VAL_OBLIGATORIO
   IngresarTotal     As Integer        'Si/no
   TieneNumDocHasta  As Integer        'VAL_NOPERMITIDO, VAL_OPCIONAL y VAL_OBLIGATORIO
   TieneCantBoletas  As Integer        'VAL_NOPERMITIDO, VAL_OPCIONAL y VAL_OBLIGATORIO
   ExigeRUT          As Boolean
   EsRebaja          As Boolean
   DocBoletas        As Boolean        'para indicar si va al libro de Ventas con Boletas
   TipoDocLAU        As Integer
   DocImpExp         As Boolean
   CodDocSII         As String
   CodDocDTESII      As String
   AceptaPropIVA     As Boolean
End Type

Type TipoLib_t
    id                As Long
    Nombre            As String
End Type
Public gTipoLibCod() As String   'LIBVENTAS, LIBCOMPRAS, LIBRETEN, LIBREMU
Public gTipoLib() As String      '"Libro de Compras", "Libro de Ventas", "Libro de Retenciones", "Libro de Remuneraciones"
Public gTipoLibNew() As TipoLib_t
Public gTratamiento() As TipoLib_t
Public gTipoDocLib() As String   'Matriz con Tipos de documentos, dependiendo del tipo de libro (Factura, Factura Exenta, Boleta, Nota de Crédito, etc.)
Public gTipoDoc() As TipoDoc_t   'Arreglo de datos de tipos de docs

Public Const MAX_TIPODOCLIB = 30

'Privilegios
Public Const PRV_ADM_SIS = &H1&           ' Administrar Sistema
Public Const PRV_CFG_EMP = &H2&           ' Configurar empresa (V. Config, comp tipo)
Public Const PRV_ADM_EMPRESA = &H4&       ' Administrar periodos contables (crear, abrir, cerrar)
Public Const PRV_ADM_CTAS = &H8&          ' Administrar Plan Cuentas y def. cuentas básicas
Public Const PRV_ING_COMP = &H10&         ' Ingresar Comprobantes
Public Const PRV_ADM_COMP = &H20&         ' Administrar Comprobantes (anular, eliminar, ...)
Public Const PRV_ING_DOCS = &H40&         ' Ingresar Docs
Public Const PRV_ADM_DOCS = &H80&         ' Administrar Docs (cetralizar, pago automático, ...)
Public Const PRV_ADM_DEF = &H100&         ' Administrar Entidades, Áreas de Negocio, Centros de Gestión
Public Const PRV_VER_INFO = &H200&        ' Ver informes, reportes y libros
Public Const PRV_IMP_LIBOF = &H400&       ' Imprimir Libros Oficiales
Public Const PRV_ADM_TIMB = &H800&        ' Administrar folios timbraje
Public Const PRV_ADM_CONCIL = &H1000&     ' Realizar conciliación bancaria
Public Const PRV_ADM_ACTFIJOS = &H2000&   ' Administrar Activos Fijos

Public Const LAST_PRV = PRV_ADM_ACTFIJOS

'Strings con nombre para cada privilegio
Public gPrivilegios() As String


'Cuentas Razones Financieras
Public Const CTA_NUMERADOR = 1
Public Const CTA_DENOMINADOR = 2

'Razones Financieras
Type TipoRazFin_t
   id As Integer
   Nombre As String
End Type

Global gTipoRazFin() As TipoRazFin_t


'2850275
Dim lPathlDbRemu As String
Dim lEsLPRemu As Boolean
Dim lRemuSQLServer As Boolean

'fin 2850275

'2860036
Public gMembrete As Membrete_t

Type Membrete_t
   TxtTitMembrete1 As String
   TxtTitMembrete2 As String
   TxtTexto1 As String
   TxtTexto2 As String
End Type

'2860036

'2861570

Public gEmail As email_t

Type email_t
   smtp     As String
   puerto   As String
   Cuenta    As String
   contraseña    As String
   to    As String
   From As String
   Subject As String
   Body As String
   adjunto As String
End Type
'2861570

'3289932
Public vMontoBaseImpoIngreso As Double
Public vMontoBaseImpoEgreso As Double
'3289932



Public Function SetupEmpSeparadas()    'EMP_SEP
   Dim Q1 As String
   Dim Rs As Recordset

   gEmprSeparadas = True
   
   Q1 = "SELECT Codigo FROM Param WHERE Tipo=" & P_EMPSEPARADAS
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then   'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & P_EMPSEPARADAS & ", " & Abs(gEmprSeparadas) & ",'Empresas Separadas')"
      Call ExecSQL(DbMain, Q1)
      
   Else
      gEmprSeparadas = Val(vFld(Rs("Codigo")))
   
   End If
   
   Call CloseRs(Rs)
   
End Function
' 2699584 3.4 TEMA 2
Public Function SetupRegimenEmpFuente()    'REGEMPREFUE
   Dim Q1 As String
   Dim Rs As Recordset
   
   SetupRegimenEmpFuente = True
   
   If Not ExisteParam("REGEMPREFUE", 1) Then  'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'REGEMPREFUE', 1 ,'Art. 14 Letra A')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   If Not ExisteParam("REGEMPREFUE", 2) Then 'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'REGEMPREFUE', 2 ,'Art. 14 Letra D, n°3')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   If Not ExisteParam("REGEMPREFUE", 3) Then  'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'REGEMPREFUE', 3 ,'Art. 14 Letra D, n°8')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   If Not ExisteParam("REGEMPREFUE", 4) Then  'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'REGEMPREFUE', 4 ,'Art. 14 Letra B, n°1')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   If Not ExisteParam("REGEMPREFUE", 5) Then  'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'REGEMPREFUE', 5 ,'Art. 14 Letra B, n°2')"
      Call ExecSQL(DbMain, Q1)
   End If

   
End Function

Public Function SetupContabilizacion()    'REGEMPREFUE
   Dim Q1 As String
   Dim Rs As Recordset
   
   SetupContabilizacion = True
   
   
   If Not ExisteParam("CONTABILIZA", 1) Then 'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'CONTABILIZA', 1 ,'Menor Activo')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   If Not ExisteParam("CONTABILIZA", 2) Then  'no está
      Q1 = "INSERT INTO Param(Tipo, Codigo, Valor) VALUES("
      Q1 = Q1 & "'CONTABILIZA', 2 ,'Ingreso Contable')"
      Call ExecSQL(DbMain, Q1)
   End If

   
End Function

Public Function ExisteParam(Tipo As String, Cod As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   ExisteParam = False
   
   Q1 = "SELECT Valor FROM Param WHERE Tipo = '" & Tipo & "' AND codigo = " & Cod
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
            ExisteParam = True
      End If
      
      Call CloseRs(Rs)
End Function
' FIN 2699584 3.4 TEMA 2

Public Function ChkPriv(Priv As Long) As Boolean
   ChkPriv = ((Priv And gUsuario.Priv) <> 0)
   'ChkPriv = True 'por ahora
End Function


Public Function FmtStRut(ByVal Rut As String, Optional ByVal bForceRUT As Boolean = 1) As String
   Dim cRut As String
   
   If gValidRut And bForceRUT Then
      cRut = Format(Rut, "00-000-000")
      FmtStRut = ReplaceStr(cRut, "-", ".") & "-" & DV_Rut(Val(Rut))
   Else
      FmtStRut = Rut
   End If
   
End Function
' *** PAM 23 Jun 2006
' Verifica que no se haya copiado un archivo a otro año o a otro RUT
#If DATACON = 1 Then
Public Function ChkDbInfo(Db As Database, ByVal Rut As String, ByVal Ano As Integer, ByVal IdEmpresa As Long) As Boolean
   Dim Q1 As String, Rs As Recordset
   
   ChkDbInfo = True
   
   If Not gEmprSeparadas Then
      Exit Function
   End If
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo = " & TPE_DBINFO
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano     'esto no importa en empresas separadas
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(Db, Q1)
   
   If Rs.EOF Then ' es nueva o no se había agregado
      Call CloseRs(Rs)

      ' Guardo datos que identifican esta base
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor)"
      Q1 = Q1 & " VALUES(" & TPE_DBINFO & ", 1, '" & Rut & "')"
      Call ExecSQL(Db, Q1)
         
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor)"
      Q1 = Q1 & " VALUES(" & TPE_DBINFO & ", 2, '" & Ano & "')"
      Call ExecSQL(Db, Q1)
         
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor)"
      Q1 = Q1 & " VALUES(" & TPE_DBINFO & ", 3, '" & IdEmpresa & "')"
      Call ExecSQL(Db, Q1)
      
   Else
      
      Do Until Rs.EOF
         
         Select Case vFld(Rs("Codigo"))
         
            Case 1: ' RUT
               If Rut <> vFld(Rs("Valor")) Then
                  MsgBox1 "Esta base de datos es del RUT " & vFld(Rs("Valor")) & " y no corresponde al RUT de la Empresa seleccionada.", vbCritical
                  ChkDbInfo = False
               End If
            
            Case 2: ' Ano
               If Ano <> Val(vFld(Rs("Valor"))) Then
                  MsgBox1 "Esta base de datos es del año " & Val(vFld(Rs("Valor"))) & " y no corresponde al Año seleccionado.", vbCritical
                  ChkDbInfo = False
               End If
            
            Case 3: ' idEmpresa
               If IdEmpresa <> Val(vFld(Rs("Valor"))) Then   'permitimos esto para el caso en que hay que reconstruir la LPContab
               
                  If ChkPriv(PRV_ADM_EMPRESA) Then
               
                     If MsgBox1("ATENCIÓN" & vbCrLf & "Esta base de datos no corresponde a la Empresa seleccionada." & vbCrLf & vbCrLf & "¿ Desea continuar bajo su responsabilidad ?", vbCritical Or vbYesNo) = vbYes Then
                        Call AddLog("Cambia idEmpresa=" & IdEmpresa & " para " & DbMain.Name)
                        Q1 = "UPDATE ParamEmpresa SET Valor = " & IdEmpresa & " WHERE Tipo = " & TPE_DBINFO & " AND Codigo = 3"
                        Call ExecSQL(DbMain, Q1)
                        Q1 = "UPDATE Empresa SET id = " & IdEmpresa    'es un solo registro en esta tabla, por empresa año
                        Call ExecSQL(DbMain, Q1)
                        
                        Call UpdateIdEmpAno(IdEmpresa, gEmpresa.Ano)
                     Else
                        ChkDbInfo = False
                     End If
                  Else
                     MsgBox1 "Esta base de datos no corresponde a la Empresa seleccionada.", vbCritical
                     ChkDbInfo = False
                  End If
               End If
            
         End Select
      
         Rs.MoveNext
      Loop
      
      Call CloseRs(Rs)
      
   End If
      
End Function
Public Function UpdateIdEmpAno(ByVal IdEmpresa As Long, Ano As Integer)
   Dim Q1 As String

   
   Q1 = "UPDATE ActFijoCompsFicha SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE ActFijoFicha SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE AFComponentes SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE AFGrupos SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE AjustesExtLibCaja SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE AreaNegocio SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE AsistImpPrimCat SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE BaseImponible14Ter SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Cartola SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CentroCosto SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Colores SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Comprobante SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Contactos SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CT_Comprobante SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CT_MovComprobante SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CtasAjustesExCont SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Cuentas SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CuentasBasicas SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE CuentasRazon SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE DetCapPropioSimpl SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE DetCartola SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE DetSaldosAp SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE DocCuotas SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Documento SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Empresa SET Id = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Entidades SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Glosas SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE ImpAdic SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE InfoAnualDJ1847 SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE LibroCaja SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE LockAction SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE LogComprobantes SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE LogImpreso SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE MovActivoFijo SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE MovComprobante SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE MovDocumento SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Notas SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE ParamEmpresa SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE ParamRazon SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE PropIVA_TotMensual SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Socios SET IdEmpresa = " & IdEmpresa & ", Ano = " & Ano
   Call ExecSQL(DbMain, Q1, False)
   Q1 = "UPDATE Sucursales SET IdEmpresa = " & IdEmpresa
   Call ExecSQL(DbMain, Q1, False)

End Function

#End If ' DATACON

Public Function CrearMdbVacia(ByVal Ano As Integer, ByVal RutMdb As String, Optional ByVal EsFact As Boolean = False) As Boolean
   Dim Rc As Integer
#If DATACON = 1 Then
   
   If Not gEmprSeparadas Then
      CrearMdbVacia = True
      Exit Function
   End If
   
   CrearMdbVacia = False
      
   On Error Resume Next
   
   If ExistFile(gDbPath & "\Empresas\" & Ano & "\" & RutMdb) = False Then
      'If MsgBox1("¡ADVERTENCIA!, no existe información de la empresa para este año. ¿Desea crearla?", vbYesNo Or vbDefaultButton1 Or vbQuestion) <> vbYes Then
      '   Exit Function
      'End If
            
      If Not ExistFile(gDbPath & "\" & BD_VACIA) Then
         MsgBox1 "No se encontró el archivo """ & gDbPath & "\" & BD_VACIA & """." & vbCrLf & "Por favor, contacte a personal de soporte del sistema.", vbExclamation + vbOKOnly
         Exit Function
      End If
               
      Rc = MkDirect(gDbPath & "\Empresas\" & Ano)
      If Rc <> 0 And Rc <> ERR_EXIST Then
         MsgBox1 "¡ADVERTENCIA!, no se podrá crear la empresa, porque no es posible crear el directorio ..\Empresas\" & Ano & " bajo el directorio ..\Datos.", vbExclamation
         Exit Function
      End If
      
      ERR.Clear
      
      Call FileCopy(gDbPath & "\" & BD_VACIA, gDbPath & "\Empresas\" & Ano & "\" & RutMdb)
   
      If ERR = 75 Then
         MsgBox1 "¡ADVERTENCIA!, no se podrá crear la empresa, porque no se ha encontrado " & BD_VACIA & " en el directorio " & gDbPath & vbCrLf & vbCrLf & "Verifique si el archivo existe. Si no es así, búsquelo en el CD de instalación del sistema o bien comuníquese con soporte.", vbExclamation
         Exit Function
      
      ElseIf ERR = 76 Then
         MsgBox1 "¡ADVERTENCIA!, no se podrá crear la empresa, porque no existe el directorio ..\Empresas bajo el directorio ..\Datos.", vbExclamation
         Exit Function
      
      ElseIf ERR <> 0 Then
         MsgBox1 "Error al crear la base de datos de la empresa en el directorio " & gDbPath & "\Empresas\" & Ano & "." & vbCrLf & vbCrLf & Error, vbExclamation
         Exit Function
         
      End If
     
   End If
         
   On Error GoTo 0
   
   CrearMdbVacia = True
   
End Function
#End If

#If DATACON = 1 Then
Public Function OpenDb(Db As Database, ByVal DbName As String) As Boolean
   Dim ConnStr As String
   
   OpenDb = True
   
   ERR.Clear
   
   On Error Resume Next
   
   Set Db = OpenDatabase(DbName, False, False, ConnStr)  'modo no exclusivo
   
   If ERR = 3356 Then
      MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
      OpenDb = False
   End If
   
   If (ERR Or Db Is Nothing) And ERR <> 3356 Then
      MsgBox "Error " & ERR & ", " & Error & NL & DbName, vbExclamation
      OpenDb = False
   End If

   If OpenDb = True Then
      DbMainDate = GetDbNow(DbMain)

      gDbType = SqlType(Db)
   End If
   
   On Error GoTo 0

#End If
End Function


Public Function OpenDbAdm(Optional ByVal BD_Name As String = "") As Integer
#If DATACON = 1 Then
   Dim DbName As String
   Dim Buf As String, Rs As Recordset, SqlErr As String
   
   On Error Resume Next
   
   OpenDbAdm = True
   
   If BD_Name = "" Then
      BD_Name = BD_COMUN
   End If
          
   DbName = gDbPath & "\" & BD_Name
   Call AddDebug("OpenDbAdm: DbName=[" & DbName & "]")
      
   '2868088
   Call SetDbSecurity(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, gComunConnStr)
   'Call SetDbSecurity(DbName, PASSW_LEXCONT_NEW, gCfgFile, SG_SEGCFG, gComunConnStr)
    'FIN 2868088
   Call AddDebug("OpenDbAdm: de SetDbSecurity")

   If Not (DbMain Is Nothing) Then
      Call CloseDb(DbMain)
   End If
      
   ERR.Clear
   'Set DbMain = OpenDatabase(DbName, True, False, ConnStr) ' MODO EXCLUSIVO
   Set DbMain = OpenDatabase(DbName, False, False, gComunConnStr)
   
   If DbMain Is Nothing Then
      SqlErr = " Error " & ERR & ", '" & Error & "'"
      Buf = "Falló OpenDB: [" & DbName & "] ConnStr=" & (gComunConnStr <> "") & ", " & SqlErr
      Call AddLog(Buf)
   End If
   
'   gComunConnStr = Mid(gComunConnStr, 2)  'sin el ; del principio  FCA: 2 feb 2016 se comenta esta línea
   
   If ERR = 3356 Then
      MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
      OpenDbAdm = False
   End If
   
   If ERR = 3343 Then
      Call AddLog("A RepairDB")
      MsgBox1 "Se ha detectado fallo en la base de datos " & BD_Name & ", se tratará de reparar. Intente ingresar nuevamente.", vbExclamation
      Call RepairDb(DbName)
      OpenDbAdm = False
   End If
   
   If (ERR Or DbMain Is Nothing) And ERR <> 3356 And ERR <> 3343 Then
      MsgBox SqlErr & vbCrLf & "'" & DbName & "'", vbExclamation
      OpenDbAdm = False
   End If
      
   If OpenDbAdm = True Then
      gDbType = SqlType(DbMain)
      DbMainDate = GetDbNow(DbMain)
      Call AddLog("Versión DB: " & SqlVersion(DbMain))
   End If
   
   Call AddDebug("OpenDbAdm: fin " & OpenDbAdm)
   
   On Error GoTo 0
   
#End If
End Function
Public Function OpenDbEmp(Optional ByVal Rut As String = "", Optional ByVal Ano As Integer = 0) As Integer
#If DATACON = 1 Then
   Dim DbName As String
   Dim Passw As String, SqlErr As String
   Dim fType As String
   
   On Error Resume Next
   
   OpenDbEmp = True
   
   If Ano > 0 Then
      If Rut <> "" Then
         DbName = gDbPath & "\Empresas\" & Ano & "\" & Rut & ".mdb"
      Else
         DbName = gDbPath & "\Empresas\" & Ano & "\" & gEmpresa.Rut & ".mdb"
      End If
      
   ElseIf Rut <> "" Then
      DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & Rut & ".mdb"
   Else
      DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & gEmpresa.Rut & ".mdb"
   End If

'   If Rut <> "" Then
'      Passw = PASSW_PREFIX & Rut
'   Else
'      Passw = PASSW_PREFIX & gEmpresa.Rut
'   End If
   
   '2868088
   If Rut <> "" Then
      Passw = PASSW_PREFIX & Rut
   Else
      Passw = PASSW_PREFIX & gEmpresa.Rut
   End If
   'FIN 2868088
   

   
   Call AddLog("OpenDbEmp: DbName:[" & DbName & "]", 2)
   
   Call SetDbSecurity(DbName, Passw, gCfgFile, SG_SEGCFG, gEmpresa.ConnStr)

   If Not (DbMain Is Nothing) Then
      Call CloseDb(DbMain)
   End If
   
   ERR.Clear
   'Set DbMain = OpenDatabase(DbName, True, False, ConnStr) ' MODO EXCLUSIVO
   Set DbMain = OpenDatabase(DbName, False, False, gEmpresa.ConnStr)
'   gEmpresa.ConnStr = Mid(gEmpresa.ConnStr, 2) 'sin el ; del principio    FCA: 2 feb 2016 se comenta esta línea
   
   If ERR Then
      SqlErr = "Error " & ERR & ", '" & Error & "'"
   
      If ERR = 3356 Then
         MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
         OpenDbEmp = False
      End If
   
   End If
   
   If (ERR Or DbMain Is Nothing) And ERR <> 3356 Then
      MsgBox "Error al abrir la base." & vbCrLf & SqlErr & vbCrLf & DbName, vbExclamation
      OpenDbEmp = False
   End If
   
   Call ChkDbSize(DbMain, 200 * 1024) ' 200 MB
   
   If OpenDbEmp = True Then
      DbMainDate = GetDbNow(DbMain)

      gDbType = SqlType(DbMain)
   End If
   
   On Error GoTo 0
   
   Call AddLog("OpenDbEmp: fin OK", 2)

#End If
End Function
'Idem OpenDBEmp pero no lo asigna a DbMain sino que retorna la variable Database
#If DATACON = 1 Then
Public Function OpenDbEmp2(DbEmp As Database, Optional ByVal Rut As String = "", Optional ByVal Ano As Integer = 0) As Integer

   Dim DbName As String
   Dim Passw As String, SqlErr As String
   
   On Error Resume Next
   
   OpenDbEmp2 = True
          
   If Ano > 0 Then
      If Rut <> "" Then
         DbName = gDbPath & "\Empresas\" & Ano & "\" & Rut & ".mdb"
      Else
         DbName = gDbPath & "\Empresas\" & Ano & "\" & gEmpresa.Rut & ".mdb"
      End If
      
   ElseIf Rut <> "" Then
      DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & Rut & ".mdb"
   Else
      DbName = gDbPath & "\Empresas\" & gEmpresa.Ano & "\" & gEmpresa.Rut & ".mdb"
   End If

'   If Rut <> "" Then
'      Passw = PASSW_PREFIX & Rut
'   Else
'      Passw = PASSW_PREFIX & gEmpresa.Rut
'   End If
   
    '2868088
   If Rut <> "" Then
      Passw = PASSW_PREFIX & Rut
   Else
      Passw = PASSW_PREFIX & gEmpresa.Rut
   End If
   'FIN 2868088
   
   
   Call AddLog("OpenDbEmp2: DbName:[" & DbName & "]", 2)
   
   Call SetDbSecurity(DbName, Passw, gCfgFile, SG_SEGCFG, gEmpresa.ConnStr)

   ERR.Clear
   'Set DbEmp = OpenDatabase(DbName, True, False, ConnStr) ' MODO EXCLUSIVO
   Set DbEmp = OpenDatabase(DbName, False, False, gEmpresa.ConnStr)
   'gEmpresa.ConnStr = Mid(gEmpresa.ConnStr, 2) 'sin el ; del principio   FCA: 2 feb 2016 se comenta esta línea
   
   If ERR Then
      SqlErr = "Error " & ERR & ", '" & Error & "'"
   
      If ERR = 3356 Then
         MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
         OpenDbEmp2 = False
      End If
   
   End If
   
   If (ERR Or DbEmp Is Nothing) And ERR <> 3356 Then
      MsgBox SqlErr & vbCrLf & DbName, vbExclamation
      OpenDbEmp2 = False
   End If
   
   Call ChkDbSize(DbMain, 200 * 1024) ' 200 MB
   
   On Error GoTo 0
   
   Call AddLog("OpenDbEmp2: fin OK", 2)

End Function
#End If


#If DATACON = 2 Then

' Para MsSql Server
' Verificar SHOW VARIABLES LIKE 'lower_case_table_names'  que sea 1 o 2
Function OpenMsSql() As Boolean
   Dim Rc As Integer, SqlPort As Long, Usr As String, Psw As String, i As Integer
   Dim ConnStr As String, Host As String, UsrPsw As String, DbName As String
   Dim sErr1 As Long, sError1 As String, Encript As Boolean, bHost As Boolean
   Const SqlSect As String = "MS Sql"

   On Error Resume Next
   
   OpenMsSql = False

   If Not DbMain Is Nothing Then
      DbMain.Close
      Set DbMain = Nothing
   End If
    
   Host = Trim(GetIniString(gCfgFile, SqlSect, "Host", ""))

   If Host = "" Then ' 15 nov 2019: se pide ubicación del host
      Host = Trim(InputBox("Ingrese la ubicación del servidor y la instancia de la base de datos 'LpContab', por ejemplo: Servidor\SqlExpress", App.Title))
   
      If Host = "" Then
         End
      End If
      
      bHost = True ' para que lo grabe
   End If
      
   SqlPort = Val(GetIniString(gCfgFile, SqlSect, "Port", "1433"))
      
   Debug.Print "Db LpContab=" & FwEncrypt1("               LpContab             ", 56516)
   DbName = GetIniString(gCfgFile, SqlSect, "DB", FwDecrypt1("55914E8C4B8B4C8E51955A2067AFB6BFB6B4B2DD5C805123764A9F754C247D57328E6B49", 56516))

   Debug.Print "User LpContab=" & FwEncrypt1("               LpContab             ", 56516)
   Usr = GetIniString(gCfgFile, SqlSect, "User", FwDecrypt1("55914E8C4B8B4C8E51955A2067AFB6BFB6B4B2DD5C805123764A9F754C247D57328E6B49", 56516))

   Debug.Print "Hola Psw=" & FwEncrypt1("     " & DbName & "   #" & "      hola       ", 731982) ' ojo con el #
   Debug.Print "Oficial Psw=" & FwEncrypt1("     " & DbName & "   #" & "     _F&].[r94%.        ", 731982) ' ojo con el #
   
   Psw = GetIniString(gCfgFile, SqlSect, "Psw")
   If Psw = "" Then
      Psw = GetIniString(gCfgFile, SqlSect, "Pswk")
      Psw = Trim(FwDecrypt1(Psw, 731982))
      
      If Psw = "" Then
         Psw = Trim(InputBox("Ingrese la clave del usuario '" & Usr & "' de la base de datos.", App.Title))
         If Psw = "" Then
            End
         End If
         Encript = True
      Else
         i = InStr(Psw, "#")
         Psw = Trim(Mid(Psw, i + 1))
         Encript = False
      End If
   Else
      Encript = True
   End If
   
   UsrPsw = "U" & "ID=" & Usr & ";P" & "WD=" & Psw & ";"
      
   ConnStr = "Driver={SQL Server};Server=" & Host & ";MARS_Connection=yes;MultipleActiveResultSets=True;Database=" & DbName & ";" ' 2 abr 2018
        
   On Error Resume Next

'   Set DbMain = OpenDatabase("", False, False, ConnStr & Usr)
      
   Set DbMain = New ADODB.Connection
   DbMain.ConnectionString = ConnStr & UsrPsw
   '3025162
   DbMain.CursorLocation = adUseClient
   DbMain.CommandTimeout = 0
   '3025162
   DbMain.Open
      
   If ERR Then
      sErr1 = ERR.Number
      sError1 = ERR.Description

      If ERR <> 3059 Then
         MsgBox1 "Problemas para conectarse a la base de datos: verifique la ubicación del servidor y la clave." & vbCrLf & vbCrLf & "Revise el archivo LPContab.cfg en la carpeta de la aplicación.", vbExclamation
'         MsgBox1 "Error " & Err & ", " & Error & vbLf & ConnStr, vbCritical
         Call AddLog("Error " & sErr1 & ", " & sError1 & ", [" & ConnStr & "]")
      End If
      
      Set DbMain = Nothing
      
      End
      Exit Function
   Else
      OpenMsSql = True
      
      If Psw = "" Then
         Psw = GetConnectInfo(DbMain, "PWD")
         UsrPsw = "User=" & Usr & ";PWD=" & Psw & ";"
         Encript = True
      End If
      
'      gConnStr = ConnStr & UsrPsw   ' Para la exportación
      
      DbMainDate = GetDbNow(DbMain)
      gDbType = SQL_SERVER

      If bHost Then ' 15 nov 2019: se graba el Host
         Call SetIniString(gCfgFile, SqlSect, "Host", Host)
      End If

      Call AddLog("Versión DB: " & SqlVersion(DbMain))

      If Encript Then
         Call SetIniString(gCfgFile, SqlSect, "Pswk", FwEncrypt1(Space(11) & DbName & Space(5) & "#" & Space(7) & Psw & Space(13), 731982))
         Call SetIniString(gCfgFile, SqlSect, "Psw", vbNullString)
      End If
   End If

   On Error GoTo 0

End Function

#End If ' DATACON

Public Function ReadComun()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim CurYear As Long
   
   Call AddLog("ReadComun: entramos", 1)
   
   Q1 = "SELECT idMoneda, Descrip, Simbolo, DecInf, DecVenta, Caracteristica, EsFijo FROM Monedas ORDER BY IdMoneda"
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 0
   ReDim gMonedas(0)
   
   Do While Rs.EOF = False
       
      ReDim Preserve gMonedas(i)
      gMonedas(i).id = vFld(Rs("IdMoneda"))
      gMonedas(i).Descrip = vFld(Rs("Descrip"), True)
      gMonedas(i).Simbolo = vFld(Rs("Simbolo"), True)
      gMonedas(i).DecInf = vFld(Rs("DecInf"))
      gMonedas(i).DecVenta = vFld(Rs("DecVenta"))
      gMonedas(i).Caract = vFld(Rs("Caracteristica"))
      gMonedas(i).EsFijo = vFld(Rs("EsFijo"))
      
      If gMonedas(i).DecInf > 0 Then
         gMonedas(i).FormatInf = NUMFMT & "." & String(gMonedas(i).DecInf, "0")
      Else
         gMonedas(i).FormatInf = NUMFMT
      End If
      
      If gMonedas(i).DecVenta > 0 Then
         gMonedas(i).FormatVenta = NUMFMT & "." & String(gMonedas(i).DecVenta, "0")
      Else
         gMonedas(i).FormatVenta = NUMFMT
      End If
      
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call AddLog("ReadComun: ya leimos monedas", 1)
  
   CurYear = DateSerial(gEmpresa.Ano, 1, 1)
   
'   Q1 = "SELECT Porcentaje FROM Impuestos WHERE Impuesto='IMPRET'"
'   Set Rs = OpenRs(DbMain, Q1)
   
'   Q1 = "SELECT Porcentaje, FechaDesde FROM Impuestos WHERE Impuesto='IMPNAC' AND (FechaDesde IS NULL OR FechaDesde <= " & CurYear & ") ORDER BY FechaDesde desc"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   gImpRet(IMPRET_NAC) = 0.1
'   If Rs.EOF = False Then
'      gImpRet(IMPRET_NAC) = vFld(Rs(0))
'   End If
'   Call CloseRs(Rs)
   
   gImpRet(IMPRET_NAC) = ImpBolHono(CurYear)

   Q1 = "SELECT Porcentaje, FechaDesde FROM Impuestos WHERE Impuesto='IMPEXT' AND (FechaDesde IS NULL OR FechaDesde <= " & CurYear & ") ORDER BY FechaDesde desc"
   Set Rs = OpenRs(DbMain, Q1)
   
   gImpRet(IMPRET_EXT) = 0.2
   If Rs.EOF = False Then
      gImpRet(IMPRET_EXT) = vFld(Rs(0))
   End If
   Call CloseRs(Rs)

   gImpRet(IMPRET_OTRO) = 0      'se calcula a partir del bruto y el impuesto
         
   Call AddLog("ReadComun: Leimos impuestos y nos vamos", 1)

End Function

Public Sub EnableForm(Frm As Form, bool As Boolean)
   Dim i As Integer
   Dim Name As String
   
   'Primero Desahabilito todo
   Call EnableForm0(Frm, bool)

   For i = 0 To Frm.Controls.Count - 1

      Name = UCase(Frm.Controls(i).Name)
      
      If Name = "BT_COPYEXCEL" Or Name = "BT_PRINT" Or Name = "BT_CLOSE" Or Name = "BT_CERRAR" Or Name = "BT_CANCEL" Or Name = "BT_PREVIEW" Or Name = "BT_CALENDAR" Or Name = "BT_CALC" Or Name = "BT_CONVMONEDA" Or Name = "BT_SUM" Then
         Frm.Controls(i).Enabled = True
         
      End If
      
   Next i
   
End Sub


Public Function GenCompAperSinMovs(ByVal NumCompAper As Long, ByVal IdEmpresa As Long, ByVal Ano As Integer, IdCompAperTrib As Long) As Long
   Dim Rs As Recordset
   Dim IdCompAper As Long
   Dim Q1 As String
   Dim FldArray(11) As AdvTbAddNew_t
   
   'generamos comprobante de apertura tributario sin movimeintos
   
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA
   Q1 = Q1 & " AND idEmpresa=" & IdEmpresa & " AND Ano=" & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then   'ya existe comprobante de apertura
      GenCompAperSinMovs = vFld(Rs("IdComp"))
      Call CloseRs(Rs)
      Exit Function
   End If
   
   Call CloseRs(Rs)
      
   
   FldArray(0).FldName = "IdUsuario"
   FldArray(0).FldValue = gUsuario.IdUsuario
   FldArray(0).FldIsNum = True
   
   FldArray(1).FldName = "FechaCreacion"
   FldArray(1).FldValue = CLng(Int(Now))
   FldArray(1).FldIsNum = True
         
   FldArray(2).FldName = "IdEmpresa"
   FldArray(2).FldValue = IdEmpresa
   FldArray(2).FldIsNum = True
               
   FldArray(3).FldName = "Ano"
   FldArray(3).FldValue = Ano
   FldArray(3).FldIsNum = True
   
   FldArray(4).FldName = "Correlativo"
   FldArray(4).FldValue = NumCompAper
   FldArray(4).FldIsNum = True
   
   FldArray(5).FldName = "Tipo"
   FldArray(5).FldValue = TC_APERTURA
   FldArray(5).FldIsNum = True
   
   FldArray(6).FldName = "Fecha"
   FldArray(6).FldValue = CLng(DateSerial(Ano, 1, 1))
   FldArray(6).FldIsNum = True
   
   FldArray(7).FldName = "Glosa"
   FldArray(7).FldValue = "Comprobante de Apertura"
   FldArray(7).FldIsNum = False
   
   FldArray(8).FldName = "Estado"
   FldArray(8).FldValue = EC_APROBADO
   FldArray(8).FldIsNum = True
   
   FldArray(9).FldName = "TipoAjuste"
   FldArray(9).FldValue = TAJUSTE_FINANCIERO
   FldArray(9).FldIsNum = True
   
   FldArray(10).FldName = "TotalDebe"
   FldArray(10).FldValue = 0
   FldArray(10).FldIsNum = True
   
   FldArray(11).FldName = "TotalHaber"
   FldArray(11).FldValue = 0
   FldArray(11).FldIsNum = True
   
   IdCompAper = AdvTbAddNewMult(DbMain, "Comprobante", "IdComp", FldArray)
   
   '3376884
    Call SeguimientoComprobantes(IdCompAper, gEmpresa.id, gEmpresa.Ano, "HyperComun.GenCompAperSinMovs", "", 1, "", gUsuario.IdUsuario, 1, 1)
    'Fin 3376884


'   Set Rs = DbMain.OpenRecordset("Comprobante")
'   Rs.AddNew
'
'   IdCompAper = Rs("IdComp")
'   Rs.Fields("Correlativo") = NumCompAper
'   Rs.Fields("Tipo") = TC_APERTURA
'   Rs.Fields("Fecha") = DateSerial(Ano, 1, 1)
'   Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'   Rs.Fields("FechaCreacion") = CLng(Int(Now))
'   Rs.Fields("Glosa") = "Comprobante de Apertura"
'   Rs.Fields("Estado") = EC_APROBADO
'   Rs.Fields("TipoAjuste") = TAJUSTE_FINANCIERO
'   Rs.Fields("TotalDebe") = 0
'   Rs.Fields("TotalHaber") = 0
'
'   Rs.Update
'   Rs.Close
'   Set Rs = Nothing
      
   
   'generamos comprobante de apertura tributario
   
   
'   Set Rs = DbMain.OpenRecordset("Comprobante")
'   Rs.AddNew
'
'   IdCompAperTrib = Rs("IdComp")
'   Rs.Fields("Correlativo") = NumCompAper
'   Rs.Fields("Tipo") = TC_APERTURA
'   Rs.Fields("Fecha") = DateSerial(Ano, 1, 1)
'   Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'   Rs.Fields("FechaCreacion") = CLng(Int(Now))
'   Rs.Fields("Glosa") = "Comprobante de Apertura"
'   Rs.Fields("Estado") = EC_APROBADO
'   Rs.Fields("TipoAjuste") = TAJUSTE_TRIBUTARIO
'   Rs.Fields("TotalDebe") = 0
'   Rs.Fields("TotalHaber") = 0
'
'   Rs.Update
'   Rs.Close
'   Set Rs = Nothing
   
   FldArray(9).FldName = "TipoAjuste"    'lo único que cambia es esto
   FldArray(9).FldValue = TAJUSTE_TRIBUTARIO
   FldArray(9).FldIsNum = True
   
   IdCompAperTrib = AdvTbAddNewMult(DbMain, "Comprobante", "IdComp", FldArray)
   
   '3376884
    Call SeguimientoComprobantes(IdCompAperTrib, gEmpresa.id, gEmpresa.Ano, "HyperComun.GenCompAperSinMovs", "", 1, "", gUsuario.IdUsuario, 1, 1)
    'Fin 3376884
   
   'Guardamos Id Comprobante de apertura
   Q1 = "UPDATE EmpresasAno SET IdCompAper=" & IdCompAper & ", NCompAper=" & NumCompAper
   Q1 = Q1 & ", IdCompAperTrib=" & IdCompAperTrib & ", NCompAperTrib=" & NumCompAper    'el correlativo es el mismo que el de apertura normal
   Q1 = Q1 & " WHERE idEmpresa=" & IdEmpresa & " AND Ano=" & Ano
   Call ExecSQL(DbMain, Q1)

   GenCompAperSinMovs = IdCompAper

End Function


Public Sub InsertParamEmpBas(ByVal IdEmpresa As Long, ByVal Ano As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim NoHayNiv As Boolean
      
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE (Tipo='NIVELES' AND " & GenLike(DbMain, "DIGNIV", "Tipo") & ")"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then   'no hay definición de niveles para la empresa
      NoHayNiv = True
   End If
   
   Call CloseRs(Rs)
   
   If NoHayNiv Then
   
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'NIVELES', 1, '4')"
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'DIGNIV1', 1, '1')"
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'DIGNIV2', 2, '2')"
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'DIGNIV3', 3, '2')"
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'DIGNIV4', 4, '2')"
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & "(IdEmpresa, Ano, Tipo, Codigo, Valor)"
      Q1 = Q1 & "VALUES(" & IdEmpresa & "," & Ano & ",'DIGNIV5', 5, '0')"
      Call ExecSQL(DbMain, Q1, False)
      
   End If
   
   'insertamos datos básicos a ParamEmpresa sin IdEmpresa-Ano porque son generales
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='IMP1CAT'"
   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0 "
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then   'no hay definición de impuestos
      
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('IMP1CAT',2016,'0.24', 0, 0)"
      Call ExecSQL(DbMain, Q1, False)
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('IMP1CAT',2017,'0.25', 0, 0)"
      Call ExecSQL(DbMain, Q1, False)
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('IMP1CAT',0,'0.25', 0, 0)"
      Call ExecSQL(DbMain, Q1, False)
   
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('VALORIVA',0,'0.19', 0, 0)"
      Call ExecSQL(DbMain, Q1, False)
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='PPM'"
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
      
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('PPM',0,'0', " & IdEmpresa & ", " & Ano & ")"
      Call ExecSQL(DbMain, Q1, False)

   End If
   
   Call CloseRs(Rs)

End Sub


'*** 9 MAY 2005 PAM - Recibe control DriveListBox para determinar path absoluto de unidades mapeadas
' vemos si es una unidad de Red y ubicamos su mapeo real
'Public Function GetAbsPath_old(ByVal Path As String, Drv As DriveListBox) As String
'   Dim i As Integer, j As Integer, k As Integer, Aux As String
'
'   Path = Trim(Path)
'
'   If Mid(Path, 2, 1) = ":" Then
'      For i = 0 To Drv.ListCount - 1
'         If UCase(Left(Path, 2)) = UCase(Left(Drv.List(i), 2)) Then
'            Aux = Drv.List(i)
'            j = InStr(Aux, "[")
'            k = InStr(Aux, "]")
'            If j <> 0 And k <> 0 Then
'               Aux = Mid(Aux, j + 1, k - j - 1)
'               If Left(Aux, 2) = "\\" Then
'                  Path = ReplaceStr(Path, Left(Path, 2), Aux)
'               End If
'            End If
'
'            Exit For
'         End If
'      Next i
'   End If
'
'   GetAbsPath_old = Path
'
'End Function

Public Sub SetDbPath(Drv As DriveListBox)
   Dim DbPath As String, Rc As Long
   Dim Q1 As String, Rs As Recordset
   Dim i As Integer, j As Integer, k As Integer
      
   DbPath = GetAbsPath(gDbPath, Drv)
   If DbPath <> "" And DbPath <> gDbPath Then
      Call AddDebug("1912: SetDbPath: Se cambia gDbPath de [" & gDbPath & "] por [" & DbPath & "]")
      Debug.Print "** SetDbPath: Se cambia gDbPath de [" & gDbPath & "] por [" & DbPath & "]"
      gDbPath = DbPath
   End If
   
   ' 16 mar 2012: para poder forzar a que no lea la tabla LParam (Lnk = 0)
   i = Val(GetIniString(gCfgFile, "Config", "Local", "0"))
   
   
   If gDbType = SQL_ACCESS Then
      If Left(gDbPath, 2) <> "\\" And i = 0 Then
   
         Q1 = "SELECT Valor FROM LParam WHERE Codigo=1"
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            DbPath = vFld(Rs("Valor"), True)
         Else
            DbPath = ""
         End If
         Call CloseRs(Rs)
   
         If Left(DbPath, 2) = "\\" Then
            If SameMdb(gDbPath & "\" & BD_COMUN, DbPath & "\" & BD_COMUN, True) Then
            
               Call AddDebug("SetDbPath: Se cambia gDbPath de [" & gDbPath & "] por [" & DbPath & "]")
               Debug.Print "** SetDbPath: Se cambia gDbPath de [" & gDbPath & "] por [" & DbPath & "]"
               gDbPath = DbPath
            Else
               Q1 = "UPDATE LParam SET Valor='" & ParaSQL(gDbPath) & "' WHERE Codigo=1"
               Rc = ExecSQL(DbMain, Q1)
            End If
         End If
   
      Else
         Q1 = "UPDATE LParam SET Valor='" & ParaSQL(gDbPath) & "' WHERE Codigo=1"
         Rc = ExecSQL(DbMain, Q1)
   
      End If

   End If
   Call AddDebug("SetDbPath: gDbPath= [" & gDbPath & "]", 2)

End Sub

Public Sub FillRegion(CbRegion As Control)

   Call AddItem(CbRegion, "< Ninguna >", 0)
   Call AddItem(CbRegion, "01  Tarapacá", 1)
   Call AddItem(CbRegion, "02  Antofagasta", 2)
   Call AddItem(CbRegion, "03  Atacama", 3)
   Call AddItem(CbRegion, "04  Coquimbo", 4)
   Call AddItem(CbRegion, "05  Valparaíso", 5)
   Call AddItem(CbRegion, "06  Lib. Gral. Bernardo O'Higgins", 6)
   Call AddItem(CbRegion, "07  Maule", 7)
   Call AddItem(CbRegion, "08  Bío Bío", 8)
   Call AddItem(CbRegion, "09  Araucanía", 9)
   Call AddItem(CbRegion, "10  Los Lagos", 10)
   Call AddItem(CbRegion, "11  Aysén", 11)
   Call AddItem(CbRegion, "12  Magallanes", 12)
   Call AddItem(CbRegion, "13  Metropolitana", 13)
   Call AddItem(CbRegion, "14  Los Ríos", 14)
   Call AddItem(CbRegion, "15  Arica y Parinacota", 15)
   Call AddItem(CbRegion, "16  Ñuble", 16)

End Sub
Public Sub AddLog(ByVal Msg As String, Optional ByVal Dbg As Integer = 0)
   Dim nErr As Long, sErr As String, Fd As Integer

   If gDebug < Dbg Then
      Exit Sub
   End If

   nErr = ERR.Number
   sErr = ERR.Description
   
   On Error Resume Next

   Fd = FreeFile
   Open W.AppPath & "\Log\LPC-" & Format(Now, "yyyymm") & ".log" For Append Access Write As #Fd

   Print #Fd, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & W.PcName & vbTab & gUsuario.Nombre & vbTab & gEmpresa.Rut & vbTab & gEmpresa.Ano & vbTab & Msg
   
   Close #Fd
   On Error GoTo 0

   ERR.Number = nErr
   ERR.Description = sErr

End Sub
Public Sub AddLogImp(ByVal FNameLogImp As String, ByVal fname As String, ByVal Linea As Integer, ByVal Msg As String)
   Dim Er As Integer, sErr As String, Fd As Integer

   Er = ERR
   sErr = Error
   On Error Resume Next

   Fd = FreeFile
   Open FNameLogImp For Append Access Write As #Fd

   Print #Fd, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & fname & vbTab & "Línea: " & Linea & vbTab & Msg
   
   Close #Fd
   On Error GoTo 0

   ERR = Er

End Sub
Public Function AddDebug(ByVal Msg As String, Optional ByVal Dbg As Integer = 0)

   If gDebug Then
      Debug.Print Msg
      Call AddLog(Msg, gDebug)
   End If

End Function
Public Function ChkVMant(ByVal VMant As Long) As Boolean

   ChkVMant = (VMant <= gAppCode.NivProd)

End Function

' Inscribe el PC en la tabla
Public Sub InscribPC()
   Dim Buf As String, Rc As Long, Mac As String, PC As String, Rs As Recordset, CodPc As String
   Dim i As Integer, iFree As Integer, Fnd As Boolean, n As Integer, Hoy As Long
   
   Call FwInitAppCode
      
'   Mac = GetMacAddress()
   Mac = GetMac()
   If Mac = "" Then
      Mac = "??-??-??-??"
   End If
   
   Hoy = Int(Now)
   
   PC = W.PcName
   CodPc = FwGetPcCode()
      
   Fnd = False
   iFree = -1
   n = 0
   For i = 1 To 150
      Buf = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_NOM & i), KEY_CRYP + i * 10)
      If Buf = "" Then
         If iFree = -1 Then
            iFree = i   ' guardamos el primero libre
         End If
      Else
         ' Buscamos si ya estaba
         n = n + 1
         If StrComp(Buf, PC, vbTextCompare) = 0 Then
            Buf = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_COD & i), KEY_CRYP + i * 75)
            If StrComp(Buf, CodPc, vbTextCompare) = 0 Then
               Buf = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_MAC & i), KEY_CRYP + i * 30)
               If StrComp(Buf, Mac, vbTextCompare) = 0 Then
                  Fnd = True
                  Call SetIniString(gLicFile, PC_EQUIP, PC_ULT & i, FwEncrypt1(Hoy, KEY_CRYP + i * 137))
                  
                  If FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_FEC & i), KEY_CRYP + i * 167) = "" Then
                     Call SetIniString(gLicFile, PC_EQUIP, PC_FEC & i, FwEncrypt1(Hoy, KEY_CRYP + i * 167))
                  End If
                  
                  Exit For
               End If
            End If
         End If
      End If
   Next i
   
   If Fnd = False Then
      Call SetIniString(gLicFile, PC_EQUIP, PC_NOM & iFree, FwEncrypt1(PC, KEY_CRYP + iFree * 10))
      Call SetIniString(gLicFile, PC_EQUIP, PC_MAC & iFree, FwEncrypt1(Mac, KEY_CRYP + iFree * 30))
      Call SetIniString(gLicFile, PC_EQUIP, PC_COD & iFree, FwEncrypt1(CodPc, KEY_CRYP + iFree * 75))
      Call SetIniString(gLicFile, PC_EQUIP, PC_AUT & iFree, FwEncrypt1("No", KEY_CRYP + iFree * 155))
      Call SetIniString(gLicFile, PC_EQUIP, PC_FEC & iFree, FwEncrypt1(Hoy, KEY_CRYP + iFree * 167))
      Call SetIniString(gLicFile, PC_EQUIP, PC_ULT & iFree, FwEncrypt1(Hoy, KEY_CRYP + iFree * 137))
   End If
   
End Sub

Public Function CheckInscPC() As Boolean
   Dim i As Integer
   Dim Buf As String, Rut As String
   Dim PC As String, Mac As String, Cod As String, NetCode As String, EstePc As Boolean
   Dim Pc1 As String, Mac1 As String, Cod1 As String, Aut1 As String
   Dim Nivel As Long, Chk As Boolean, nPC As Integer, nLic As Integer, bLic As Boolean
   
   On Error Resume Next
   
   CheckInscPC = False
   
   gCantLicencias = 0
   
   PC = GetComputerName()
   Cod = FwGetPcCode()
   Mac = GetMacAddress()

   If Mac = "" Then
      Mac = "??-??-??-??"
   End If
   
   NetCode = Trim(FwDecrypt1(GetIniString(gLicFile, PC_INFO, PC_NCOD & 1), KEY_CRYP + 2345))
   Nivel = Val(FwDecrypt1(GetIniString(gLicFile, PC_INFO, PC_NIV & 3), KEY_CRYP + 3147)) - 654321
   Rut = Trim(FwDecrypt1(GetIniString(gLicFile, PC_INFO, PC_RUT & 1), KEY_CRYP + 7145))
   
   Buf = GetIniString(gLicFile, PC_INFO, PC_NLIC & 3)
   If Buf <> "" Then
      bLic = True
      nLic = (Val(FwDecrypt1(Buf, KEY_CRYP + 5043)) - 735081) / 19
   Else
      bLic = False
   End If
            
   If Len(NetCode) < 5 Then ' no tiene el nuevo esquema o es demo
   
      If gAppCode.Demo = True Then
         gCantLicencias = 1 ' demo
      End If
      Exit Function
   End If
   
   gAppCode.Demo = True
   gAppCode.NivProd = Nivel
   gAppCode.Rut = Rut
   
   Chk = False
   
   ' PAM: 14 ago 2008, se deja el registro anterior
   ' Call FwUnRegister ' ya no corre el esquema anterior
      
   Buf = ""
   EstePc = False
   nPC = 0
   For i = 1 To 100
   
      Pc1 = GetIniString(gLicFile, PC_EQUIP, PC_NOM & i)
      
      If Pc1 <> "" Then
         Pc1 = FwDecrypt1(Pc1, KEY_CRYP + i * 10)
         Mac1 = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_MAC & i), KEY_CRYP + i * 30)
         Cod1 = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_COD & i), KEY_CRYP + i * 75)
         Aut1 = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_AUT & i), KEY_CRYP + i * 155)
         
         If StrComp(Aut1, "Si", vbTextCompare) = 0 Then
            nPC = nPC + 1
            Buf = Buf & "::" & Pc1 & ":" & Mac1 & ":" & Cod1 & ":"
         
            'If PC = Pc1 And Mac = Mac1 And Cod = Cod1 Then ** 10 nov 2010: No controlamos la MAC porque cambia
            If PC = Pc1 And Cod = Cod1 Then
               EstePc = True
            End If
         
         End If
      End If
      
   Next i
   
   If bLic Then
      Buf = Buf & "[" & Nivel & ":" & nPC & ":" & nLic & "]"
   Else
      Buf = Buf & "[" & Nivel & ":" & nPC & "]"
      nLic = nPC
   End If

   Buf = Buf & "RUT:" & Rut & ";"
   Buf = Buf & "PRD:" & APP_NAME & ";"
   
   Chk = (NetCode = GenCode(UCase(Buf), PC_SEED))
      
   If Chk And EstePc Then
      gAppCode.Demo = False
      gAppCode.NivProd = Nivel
      gCantLicencias = nLic
   Else
      gAppCode.Demo = True
      gAppCode.NivProd = VER_DEMO
      gCantLicencias = 1 ' demo
   End If
   
   CheckInscPC = Chk
   
End Function


Public Function GenCodigo(ByVal Info As String) As String

   Info = ReplaceStr(Info, " ", "")
   Info = ReplaceStr(Info, vbCr, "")
   Info = ReplaceStr(Info, vbLf, "")

   GenCodigo = GenCode(Info, PC_SEED)

End Function
'Si año = 0 entonces asume gEmpresa.Ano, si no el año actual
Public Sub ReadIndices(Optional ByVal Ano As Integer = 0)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim m1 As Long, M2 As Long
   Dim i As Integer, j As Integer
   Dim Mes As Integer
   Dim IdxAno As Integer
   
   If Ano = 0 Then
      IdxAno = gEmpresa.Ano
   Else
      IdxAno = Ano
   End If
   
   If IdxAno = 0 Then
      IdxAno = Year(Now)
   End If
   
   For i = 0 To 12
      gIndices(i).PuntosIPC = 0
      gIndices(i).VarIpc = 0
      gIndices(i).FactorCM = 0
   Next i
      
   m1 = DateSerial(IdxAno - 1, 12, 1)
   M2 = DateSerial(IdxAno, 12, 1)
   
   Q1 = "SELECT AnoMes, pIPC, vIPC, fCM FROM IPC WHERE AnoMes BETWEEN " & m1 & " AND " & M2 & " ORDER BY AnoMes"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Mes = DateDiff("m", m1, vFld(Rs("AnoMes")))
      If Mes >= 0 And Mes <= 12 Then    'por si las moscas
         gIndices(Mes).PuntosIPC = vFld(Rs("pIPC"))
         gIndices(Mes).VarIpc = vFld(Rs("vIPC"))
         gIndices(Mes).FactorCM = vFld(Rs("fCM"))
         
         If gEmpresa.Ano = 2019 Then         'Diconsinuidad del INE (Victor Morales, 20 ago 2019)
            If Mes = 0 Then       'Dic 2018
               gIndices(Mes).PuntosIPC = 100.64
            End If
         End If
      
      End If
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call ReadFactorActAnual(IdxAno, gFactorActAnual)
   
   
End Sub
Public Sub ReadFactorActAnual(ByVal Ano As Integer, Fact() As SII_Fact_t)
   Dim i As Integer, j As Integer
   Dim Rs As Recordset
   Dim Q1 As String

   For i = 0 To 12
      For j = 1 To 12
      
         Fact(i, j).bFact = False
         Fact(i, j).Fact = 0
         
      Next j
   Next i
   
   Q1 = "SELECT MesRow, MesCol, Factor FROM FactorActAnual WHERE Ano = " & Ano & " ORDER BY MesRow, MesCol "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Fact(vFld(Rs("MesRow")), vFld(Rs("MesCol"))).bFact = IIf(IsNull(Rs("Factor")), False, True)
      Fact(vFld(Rs("MesRow")), vFld(Rs("MesCol"))).Fact = vFld(Rs("Factor"))
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   

End Sub

'PS 19-04-2006
Public Sub CheckRcEmpAno(Ano As Integer, IdEmpresa As Long)
   Dim Q1 As String
   Dim Rs As Recordset

   Q1 = "SELECT Ano FROM EmpresasAno WHERE Ano=" & Ano & " AND idEmpresa=" & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
      'Si pasa por aquí es prq hubo una inconsistencia y se arregla con esto.
      Q1 = "INSERT INTO EmpresasAno (IdEmpresa, Ano, FApertura) VALUES ("
      Q1 = Q1 & IdEmpresa & "," & Ano & "," & CLng(Int(Now)) & ")"
      Call ExecSQL(DbMain, Q1)
   End If
   Call CloseRs(Rs)
   
End Sub
' por si hay problemas con la conexión probamos sólo un par de veces en el día
' si bMsg = true es porque lo solicitó el usuario
Public Function CheckVersion(Frm As Form, ByVal bMsg As Boolean) As Boolean
   Dim d As Double, n As Integer, H As Long

   If gChecked = False Then ' Basta con una vez
      
      d = GetIniString(gIniFile, "Config", "I1", "0")
      n = GetIniString(gIniFile, "Config", "I2", "0")
      H = GetIniString(gIniFile, "Config", "I3", "0")
      
      ' si ya intentó más de dos veces y no pudo, asumumos que si, mañana será otro día
      If bMsg = False And d = CLng(Int(Now)) And n > 2 Then ' intentamos hasta dos veces en el día
         gChecked = True
      
      Else
         If d <> CLng(Int(Now)) Or bMsg = True Then
            n = 0
            H = Hour(Now) * 60 + Minute(Now)
         End If
         
         If Hour(Now) * 60 + Minute(Now) >= H Then
            Call SetIniString(gIniFile, "Config", "I1", CLng(Int(Now)))
            Call SetIniString(gIniFile, "Config", "I2", n + 1)
            Call SetIniString(gIniFile, "Config", "I3", Int(H + 70 + (20 * Rnd))) ' prueba en 95 minutos
            
            gChecked = FwCheckVer(Frm, APP_NAME, App.Title, APP_URL, , "&r=" & gAppCode.Rut & "&cpc=" & FwGetPcCode() & "&d=" & Abs(gAppCode.Demo) & "&ver=" & W.Version & "&fver=" & Format(W.FVersion, "yyyymmdd"), bMsg)
         End If
         
      End If
      
   End If
   
   CheckVersion = gChecked
      
End Function

Public Sub ReadTipoValor()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String

   'tipos de movimientos (o de valor) para documentos de libros de compras, ventas, etc.
   ReDim gTipoValLib(10)
        
   Q1 = "SELECT idTValor, TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto, Tit1, Tit2, TipoIVARetenido"
   Q1 = Q1 & " FROM TipoValor ORDER BY TipoLib, Orden, Valor "
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 0
   Do While Rs.EOF = False
   
      If i > UBound(gTipoValLib) Then
         ReDim Preserve gTipoValLib(i + 10)
      End If
      
       '2991797-3002457
   Q2 = "SELECT IdImpAdic "
   Q2 = Q2 & " FROM ImpAdic "
   Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano & " AND IdEmpresa = " & gEmpresa.id & " And TipoValor = " & vFld(Rs("Codigo"))
   Q2 = Q2 & " And TipoLib = " & vFld(Rs("TipoLib"))

   Set Rs2 = OpenRs(DbMain, Q2)

   If Rs2.EOF = False Then

      If vFld(Rs("Tasa")) > 0 Then
        Q3 = "UPDATE ImpAdic SET Tasa = '" & IIf(gDbType = SQL_ACCESS, vFld(Rs("Tasa")), Replace(vFld(Rs("Tasa")), ",", ".")) & "' WHERE ano = " & gEmpresa.Ano & " AND IdEmpresa = " & gEmpresa.id & " And TipoValor = " & vFld(Rs("Codigo"))
        Q3 = Q3 & " And TipoLib = " & vFld(Rs("TipoLib"))
        Call ExecSQL(DbMain, Q3)
      End If
   End If

   Call CloseRs(Rs2)
   
   '2991797-3002457
         
      gTipoValLib(i).id = vFld(Rs("idTValor"))
      gTipoValLib(i).TipoLib = vFld(Rs("TipoLib"))
      gTipoValLib(i).TipoDoc = vFld(Rs("TipoDoc"))
      gTipoValLib(i).TipoValLib = vFld(Rs("Codigo"))
      gTipoValLib(i).Nombre = vFld(Rs("Valor"))
      gTipoValLib(i).Diminutivo = vFld(Rs("Diminutivo"))
      gTipoValLib(i).Atributo = vFld(Rs("Atributo"))
      gTipoValLib(i).Multiple = vFld(Rs("Multiple"))
      gTipoValLib(i).CodF29 = vFld(Rs("CodF29"))
      gTipoValLib(i).orden = vFld(Rs("Orden"))
      gTipoValLib(i).Tasa = vFld(Rs("Tasa"))
      gTipoValLib(i).TasaFija = IIf(IsNull(Rs("Tasa")), False, True)
      gTipoValLib(i).EsRecuperable = vFld(Rs("EsRecuperable"))
      gTipoValLib(i).CodSIIDTE = vFld(Rs("CodSIIDTE"))
      gTipoValLib(i).TitCompleto = IIf(vFld(Rs("TitCompleto")) <> "", vFld(Rs("TitCompleto")), vFld(Rs("Tit1")) & " " & vFld(Rs("Tit2")))
      gTipoValLib(i).Descontinuado = IIf(InStr(gTipoValLib(i).Nombre, MARCA_DESCONTINUADO) > 0, True, False)
      gTipoValLib(i).TipoIVARetenido = vFld(Rs("TipoIVARetenido"))
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   gTipoIvaIrrec(1).CodImpSII = 1
   gTipoIvaIrrec(1).Descrip = "Compras destinadas a generar operaciones no gravadas o exentas"

   gTipoIvaIrrec(1).CodImpSII = 2
   gTipoIvaIrrec(1).Descrip = "Facturas de proveedores registradas fuera de plazo"

   gTipoIvaIrrec(1).CodImpSII = 3
   gTipoIvaIrrec(1).Descrip = "Gastos rechazados"

   gTipoIvaIrrec(1).CodImpSII = 4
   gTipoIvaIrrec(1).Descrip = "Entregas gratuitas (premios, bonificaciones etc.) recibidas"

   gTipoIvaIrrec(1).CodImpSII = 9
   gTipoIvaIrrec(1).Descrip = "Otros"

End Sub

Public Sub ReadOficina()
   Dim Q1 As String, Rs As Recordset

   gOficina.Rut = ""
   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='OFICINA'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs Is Nothing Then
      MsgBox1 "La base de datos está corrupta o es muy antigua.", vbCritical
      Call CloseDb(DbMain)
      End
   End If
   
   Do Until Rs.EOF
    
      Select Case vFld(Rs("Codigo"))
         Case TOF_RUT:
            gOficina.Rut = vFld(Rs("Valor"))
            
            If Len(gAppCode.Rut) = 0 Then
               gAppCode.Rut = gOficina.Rut
            End If
      
         Case TOF_NOMBRE:
            gOficina.Nombre = vFld(Rs("Valor"))
         
      End Select
   
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)

End Sub

Public Function ExitDemo() As Boolean
   Dim Rs As Recordset, Q1 As String
   
   If W.InDesign Then
      Exit Function
   End If
      
   If gAppCode.Demo = False Then
      ExitDemo = False
      Exit Function
   End If

   ' Sólo deberían estar estos RUTs
   Q1 = "SELECT Rut FROM Empresa WHERE Rut IN ('1','2','3')"
   Set Rs = OpenRs(DbMain, Q1)
   ExitDemo = Rs.EOF
   Call CloseRs(Rs)
      
End Function
Public Function GetValMoneda(ByVal Simbolo As String, ValMoneda As Double, Optional ByVal Fecha As Long = 0, Optional ByVal ExactoFecha As Boolean = False) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   
   
   Q1 = "SELECT Fecha, Valor "
   Q1 = Q1 & " FROM Equivalencia INNER JOIN Monedas ON Equivalencia.IdMoneda = Monedas.IdMoneda"
   Q1 = Q1 & " WHERE Monedas.Simbolo = '" & Simbolo & "'"
   If Fecha > 0 And ExactoFecha Then
      Q1 = Q1 & " AND Fecha = " & Fecha
   End If
    Q1 = Q1 & " ORDER BY Fecha desc"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   GetValMoneda = False
   ValMoneda = 0
   
   Do While Not Rs.EOF
      ValMoneda = vFld(Rs("Valor"))   'tomamos el último ingresado (por la fecha)
      
      If ValMoneda > 0 Then
         GetValMoneda = True
         
         If Fecha > 0 Then
            If ExactoFecha Then    'el registro es el seleccionado
               Call CloseRs(Rs)
               Exit Function
                     
            ElseIf vFld(Rs("Fecha")) <= Fecha Then
               Call CloseRs(Rs)
               Exit Function
            End If
         Else
            Call CloseRs(Rs)
            Exit Function
         End If
         
      End If
      
      Rs.MoveNext
   Loop
   
   If ValMoneda = 0 Then
      GetValMoneda = False
   End If
   
   Call CloseRs(Rs)
      
End Function


'Public Function SeguimientoDocumento(IdDoc As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Optional ByVal Query As String = "", Optional ByVal Vigente As Integer = 0, Optional ByVal Where As String = "", Optional ByVal Usuario As String = "", Optional ByVal FIngreso As Integer = 1, Optional ByVal Ajuste As Integer = 1) As Boolean
Public Function SeguimientoDocumento(IdDoc As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Query As String, Vigente As Integer, Where As String, Usuario As Integer, FIngreso As Integer, Ajuste As Integer) As Boolean
   Dim Q1 As String, Rs As Recordset, bIns As Boolean, Pid As Long, oPid As Long, nConected As Integer
   Dim Frm As FrmDesbloquear
   Dim fechaHoy As String
   fechaHoy = "GetDate()"
   #If DATACON = 1 Then
        fechaHoy = "Now()"
   #End If
    Q1 = "INSERT INTO Tracking_Documento"
    Q1 = Q1 & "       (IdDoc,IdEmpresa,Ano,IdCompCent,IdCompPago,TipoLib,TipoDoc,NumDoc,NumDocHasta,IdEntidad,TipoEntidad,RutEntidad,NombreEntidad,FEmision"
    Q1 = Q1 & "       ,FVenc,Descrip,Estado,Exento,IdCuentaExento,Afecto,IdCuentaAfecto,IVA,IdCuentaIVA,OtroImp,IdCuentaOtroImp,Total,IdCuentaTotal,IdUsuario"
    Q1 = Q1 & "       ,FechaCreacion,FEmisionOri,CorrInterno,SaldoDoc,FExported,OldIdDoc,DTE,PorcentRetencion,TipoRetencion,MovEdited,OtrosVal,FImporF29,NumDocRef"
    Q1 = Q1 & "       ,IdCtaBanco,TipoRelEnt,IdSucursal,TotPagadoAnoAnt,FImportSuc,Giro,FacCompraRetParcial,IVAIrrecuperable,DocOtrosEnAnalitico,OldIdDocTmp,NumFiscImpr"
    Q1 = Q1 & "       ,NumInformeZ,CantBoletas,VentasAcumInfZ,IdDocAsoc,PropIVA,ValIVAIrrec,IVAInmueble,FImpFacturacion,CodSIIDTEIVAIrrec,TipoDocAsoc,IVAActFijo"
    Q1 = Q1 & "       ,EntRelacionada,NumCuotas,CompraBienRaiz,NumDocAsoc,DTEDocAsoc,IdANegCCosto,UrlDTE,CodCtaAfectoOld,CodCtaExentoOld,CodCtaTotalOld,DocOtroEsCargo"
    Q1 = Q1 & "       ,ValRet3Porc,IdCuentaRet3Porc,Tratamiento, Origen, Query, Vigente, FechaHora, FormaIngreso, Ajuste)"
    Q1 = Q1 & " SELECT IdDoc,IdEmpresa,Ano,IdCompCent,IdCompPago,TipoLib,TipoDoc,NumDoc,NumDocHasta,IdEntidad,TipoEntidad,RutEntidad,NombreEntidad,FEmision"
    Q1 = Q1 & "     ,FVenc,Descrip,Estado,Exento,IdCuentaExento,Afecto,IdCuentaAfecto,IVA,IdCuentaIVA,OtroImp,IdCuentaOtroImp,Total,IdCuentaTotal," & Usuario
    Q1 = Q1 & "     ,FechaCreacion,FEmisionOri,CorrInterno,SaldoDoc,FExported,OldIdDoc,DTE,PorcentRetencion,TipoRetencion,MovEdited,OtrosVal,FImporF29,NumDocRef"
    Q1 = Q1 & "     ,IdCtaBanco,TipoRelEnt,IdSucursal,TotPagadoAnoAnt,FImportSuc,Giro,FacCompraRetParcial,IVAIrrecuperable,DocOtrosEnAnalitico,OldIdDocTmp,NumFiscImpr"
    Q1 = Q1 & "     ,NumInformeZ,CantBoletas,VentasAcumInfZ,IdDocAsoc,PropIVA,ValIVAIrrec,IVAInmueble,FImpFacturacion,CodSIIDTEIVAIrrec,TipoDocAsoc,IVAActFijo"
    Q1 = Q1 & "     ,EntRelacionada,NumCuotas,CompraBienRaiz,NumDocAsoc,DTEDocAsoc,IdANegCCosto,UrlDTE,CodCtaAfectoOld,CodCtaExentoOld,CodCtaTotalOld,DocOtroEsCargo"
    Q1 = Q1 & "     ,ValRet3Porc,IdCuentaRet3Porc,Tratamiento, '" & Origen & "' AS Origen, '" & Replace(Query, Chr(39), "") & "' AS Query," & Vigente & " AS Vigente, " & fechaHoy & " AS FechaHora, " & FIngreso & " AS FormaIngreso, " & Ajuste & " AS Ajuste"
    Q1 = Q1 & "     FROM Documento"
    If Where <> "" Then
        Q1 = Q1 & Where & ""
    Else
        Q1 = Q1 & "     WHERE IdDoc =  " & IdDoc & ""
    End If
    Call ExecSQL(DbMain, Q1)
    
   
End Function

'Public Function SeguimientoMovDocumento(IdDoc As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Optional ByVal Query As String = "", Optional ByVal Vigente As Integer = 0, Optional ByVal Where As String = "", Optional ByVal FIngreso As Integer = 1, Optional ByVal Ajuste As Integer = 1) As Boolean
Public Function SeguimientoMovDocumento(IdDoc As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Query As String, Vigente As Integer, Where As String, FIngreso As Integer, Ajuste As Integer) As Boolean
   Dim QU1 As String, Rs As Recordset, bIns As Boolean, Pid As Long, oPid As Long, nConected As Integer
   Dim Frm As FrmDesbloquear
   
   Dim fechaHoy As String
   fechaHoy = "GetDate()"
   #If DATACON = 1 Then
        fechaHoy = "Now()"
   #End If

    
    QU1 = "INSERT INTO Tracking_MovDocumento"
    QU1 = QU1 & "       (IdMovDoc,IdEmpresa,Ano,IdDoc,IdCompCent,IdCompPago,Orden,IdCuenta,Debe,Haber,Glosa"
    QU1 = QU1 & "       ,IdTipoValLib,EsTotalDoc,IdCCosto,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld, Origen, Query, Vigente, FechaHora, FormaIngreso, Ajuste)"
    QU1 = QU1 & " SELECT IdMovDoc,IdEmpresa,Ano,IdDoc,IdCompCent,IdCompPago,Orden,IdCuenta,Debe,Haber,Glosa"
    QU1 = QU1 & "      ,IdTipoValLib,EsTotalDoc,IdCCosto,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld, '" & Origen & "' AS Origen, '" & Replace(Query, Chr(39), "") & "' AS Query," & Vigente & " AS Vigente, " & fechaHoy & " AS FechaHora, " & FIngreso & " AS FormaIngreso, " & Ajuste & " AS Ajuste"
    QU1 = QU1 & " FROM MovDocumento"
    'QU1 = QU1 & " WHERE IdMovDoc = " & IdDoc
    If Where <> "" Then
        QU1 = QU1 & Where
    Else
        QU1 = QU1 & " WHERE IdDoc = " & IdDoc & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
    End If
    Call ExecSQL(DbMain, QU1)
   
End Function

'Public Function SeguimientoComprobantes(idcomp As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Optional ByVal Query As String = "", Optional ByVal Vigente As Integer = 0, Optional ByVal Where As String = "", Optional ByVal Usuario As String = "", Optional ByVal FIngreso As Integer = 1, Optional ByVal Ajuste As Integer = 1) As Boolean
Public Function SeguimientoComprobantes(idcomp As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Query As String, Vigente As Integer, Where As String, Usuario As Integer, FIngreso As Integer, Ajuste As Integer) As Boolean
   Dim QU1 As String, Rs As Recordset, bIns As Boolean, Pid As Long, oPid As Long, nConected As Integer
   Dim Frm As FrmDesbloquear
   
   Dim fechaHoy As String
   fechaHoy = "GetDate()"
   #If DATACON = 1 Then
        fechaHoy = "Now()"
   #End If
  
    QU1 = "INSERT INTO Tracking_Comprobante"
    QU1 = QU1 & "       (IdComp,FechaHora,IdEmpresa,Ano,Correlativo,Fecha,Tipo,Estado,Glosa,TotalDebe,TotalHaber,IdUsuario"
    QU1 = QU1 & "       ,FechaCreacion,ImpResumido,EsCCMM,FechaImport,TipoAjuste,OtrosIngEg14TER,Origen,Query,Vigente, FormaIngreso, Ajuste)"
    QU1 = QU1 & " SELECT IdComp," & fechaHoy & " AS FechaHora,IdEmpresa,Ano,Correlativo,Fecha,Tipo,Estado,Glosa,TotalDebe,TotalHaber, " & Usuario
    QU1 = QU1 & "       ,FechaCreacion,ImpResumido,EsCCMM,FechaImport,TipoAjuste,OtrosIngEg14TER, '" & Origen & "' AS Origen, '" & Query & "' AS Query," & Vigente & " AS Vigente, " & FIngreso & " AS FormaIngreso, " & Ajuste & " AS Ajuste"
    QU1 = QU1 & "       FROM Comprobante"
    
    
    If Where <> "" Then
        QU1 = QU1 & Where & ""
    Else
        QU1 = QU1 & "     WHERE IdComp =  " & idcomp & ""
    End If
    Call ExecSQL(DbMain, QU1)
    
  
End Function

'Public Function SeguimientoMovComprobante(idMov As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Optional ByVal Query As String = "", Optional ByVal Vigente As Integer = 0, Optional ByVal Where As String = "", Optional ByVal FIngreso As Integer = 1, Optional ByVal Ajuste As Integer = 1) As Boolean
Public Function SeguimientoMovComprobante(idMov As Long, IdEmpresa As Long, Ano As Integer, Origen As String, Query As String, Vigente As Integer, Where As String, FIngreso As Integer, Ajuste As Integer) As Boolean
   Dim QU1 As String, Rs As Recordset, bIns As Boolean, Pid As Long, oPid As Long, nConected As Integer
   Dim Frm As FrmDesbloquear
   
   Dim fechaHoy As String
   fechaHoy = "GetDate()"
   #If DATACON = 1 Then
        fechaHoy = "Now()"
   #End If

    
    QU1 = "INSERT INTO Tracking_MovComprobante"
    QU1 = QU1 & "       (IdMov,FechaHora,IdEmpresa,Ano,IdComp,IdDoc,Orden,IdCuenta,Debe,Haber,Glosa,idCCosto"
    QU1 = QU1 & "       ,idAreaNeg,IdCartola,DeCentraliz,DePago,DeRemu,Nota,IdDocCuota,Origen,Query,Vigente, FormaIngreso, Ajuste)"
    QU1 = QU1 & " SELECT IdMov," & fechaHoy & " AS FechaHora,IdEmpresa,Ano,IdComp,IdDoc,Orden,IdCuenta,Debe,Haber,Glosa,idCCosto"
    QU1 = QU1 & "       ,idAreaNeg,IdCartola,DeCentraliz,DePago,DeRemu,Nota,IdDocCuota, '" & Origen & "' AS Origen, '" & Query & "' AS Query," & Vigente & " AS Vigente, " & FIngreso & " AS FormaIngreso, " & Ajuste & " AS Ajuste"
    QU1 = QU1 & " From MovComprobante"
    
    If Where <> "" Then
        QU1 = QU1 & Where
    Else
        QU1 = QU1 & " WHERE IdComp = " & idMov & " IdEmpresa = " & IdEmpresa & " Ano = " & Ano
    End If
    Call ExecSQL(DbMain, QU1)
   
End Function

Public Function ContRegisterPc(Optional ByVal Msg As String = "", Optional ByVal MaxLic As Integer = -1) As Boolean
   Dim Q1 As String, Rs As Recordset, bIns As Boolean, Pid As Long, oPid As Long, nConected As Integer
   Dim Frm As FrmDesbloquear
    
   ' pam: Nueva Instancia
   If gNuevaInstancia Then
      ContRegisterPc = True
      Exit Function
   End If
   
   ContRegisterPc = False
   
   oPid = -1
   Q1 = "SELECT Pid FROM PcUsr WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      bIns = True
   Else
      oPid = vFld(Rs("Pid"))
      bIns = False
   End If
   Call CloseRs(Rs)
   
   Pid = GetCurrentProcessId()
      
   Call AddLog("Register: PC='" & W.PcName & "', Usr='" & W.UserName & "', oPid=" & oPid & ", Pid=" & Pid & ", in=" & bIns)
   
   If bIns Then
CountUser:
      Q1 = "SELECT Count(*) as N FROM PcUsr"
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         nConected = vFld(Rs("N"))
      End If
      Call CloseRs(Rs)
      
      Call AddLog("Register: nCon=" & nConected & ", Max=" & MaxLic)
      
      If nConected >= MaxLic Then
         If MsgBox1("Esta conexión supera la cantidad de licencias disponibles." & vbCrLf & "¿Desea desconectar algún otro usuario antes de continuar?.", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            
            Set Frm = New FrmDesbloquear
            Frm.Show vbModal
            Set Frm = Nothing
            GoTo CountUser
         Else
            Exit Function
            
         End If
      End If
   
      Q1 = "INSERT INTO PcUsr (PC, Usr, Pid ) values ('" & ParaSQL(W.PcName) & "','" & ParaSQL(W.UserName) & "'," & Pid & " )"
      Call ExecSQL(DbMain, Q1)
      ContRegisterPc = True
      
   ElseIf Pid <> oPid Then
   
      If Msg = "" Then
         'no se permite más de un usuario en un mismo equipo para evitar que algunos usuarios multipliquen sus licencias utilizando concección remota con Terminal Server
         Msg = "El usuario de Windows '" & W.UserName & "' ya está conectado al sistema en este equipo." & vbCrLf & "¿Desea desconectarlo de la aplicación?"
      End If
   
      If oPid <> 0 Then
         If MsgBox1(Msg, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            Q1 = "UPDATE PcUsr SET Pid=" & Pid & " WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
            Call ExecSQL(DbMain, Q1)
            ContRegisterPc = True
         End If
      Else ' es la primera vez, estaba en NULL
         Q1 = "UPDATE PcUsr SET Pid=" & Pid & " WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
         Call ExecSQL(DbMain, Q1)
         ContRegisterPc = True
      End If
   Else
      ContRegisterPc = True
   End If
   
End Function
Public Sub ContUnregisterPc(Optional ByVal idFrom As Integer = 0)
   Dim Q1 As String
   
   ' pam: Nueva Instancia
   If gNuevaInstancia Then
      Exit Sub
   End If
   
'   Q1 = "DELETE * FROM PcUsr WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
   Q1 = " WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
   Call DeleteSQL(DbMain, "PcUsr", Q1)
   
   Call AddLog("Unregister: PC='" & W.PcName & "', Usr='" & W.UserName & "', From:" & idFrom & ", Pid=" & GetCurrentProcessId())

End Sub

' Verifica si sigue conectado
Public Function ContRegisteredUsr() As String
   Dim Q1 As String, Rs As Recordset, oPid As Long, Pid As Long, Usr As String
   
   oPid = -1
'   Pid = -1
   Pid = GetCurrentProcessId()
   
'   Q1 = "SELECT Usr FROM PcUsr WHERE PC='" & W.PcName & "' AND Usr='" & W.UserName & "' AND Pid<>" & GetCurrentProcessId()
   Q1 = "SELECT Usr, Pid FROM PcUsr WHERE PC='" & ParaSQL(W.PcName) & "' AND Usr='" & ParaSQL(W.UserName) & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      oPid = vFld(Rs("Pid"))
      If oPid = Pid Or gNuevaInstancia = True Then  ' pam
         Usr = ""
      Else
         Usr = vFld(Rs("Usr"))   ' se debe desconectar
      End If
   Else
      Usr = "."
   End If
   Call CloseRs(Rs)
   
   ContRegisteredUsr = Usr
   
#If DATACON = 1 Then       'Access
   Call AddLog("Registered?: PC='" & W.PcName & "', Usr='" & W.UserName & "', oPid=" & oPid & ", Pid=" & Pid & ", ni=" & gNuevaInstancia & ", usr='" & Usr & "', PathTb: '" & DbTablePath(DbMain, "PcUsr") & "'")
#End If

End Function

' pam: Nueva Instancia: dura un minuto
Public Function GenInstanceKey() As Long
   Dim Buf As String

   Buf = "$" & App.Title & "&" & Format(Now, EDATEFMT) & "#<" & W.PcName & ">#" & Format(Now, "hh:nn") & "?" & W.UserName & "%"
'   MsgBox1 "GKey=[" & Buf & "]"
   GenInstanceKey = GenClave3(Buf, PC_SEED)

End Function
'Public Function vFldADO(Fld As ADODB.Field, Optional ByVal bDeSql As Boolean = True) As Variant
'   Dim bString As Boolean, bBoolean As Boolean
'
'   bString = (Fld.Type = adChar Or Fld.Type = adVarChar Or Fld.Type = adLongVarChar Or Fld.Type = adLongVarWChar Or Fld.Type = adVarWChar Or Fld.Type = adWChar)
'   bBoolean = (Fld.Type = adBoolean)
'
'   If IsNull(Fld) Then
'
'      If bString Then
'         vFldADO = ""
'      Else
'         vFldADO = 0
'      End If
'
'   ElseIf bString Then
'
'      If bDeSql Then
'         vFldADO = DeSQL(Fld.Value)
'      Else
'         vFldADO = Fld.Value
'      End If
'
'   ElseIf bBoolean Then
'
'      vFldADO = Abs(Fld.Value)
'
'   Else
'      vFldADO = Fld.Value
'   End If
'
'End Function
Public Sub LP_FGr2Clip(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False)

'   If gAppCode.Demo Then
'      MsgBox1 "En modo Demo el sistema no permite copiar los balances y libros a Excel.", vbInformation
'      Exit Sub
'   End If
   
   Call FGr2Clip(Grid, Title, bIncludeCero)

End Sub

Public Function LP_FGr2String(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False, Optional ByVal ColOblig As Integer = -1) As String
   
'   If gAppCode.Demo Then
'      MsgBox1 "En modo Demo el sistema no permite copiar los balances y libros a Excel.", vbInformation
'      Exit Function
'   End If
   
   LP_FGr2String = FGr2String(Grid, Title, bIncludeCero, ColOblig)

End Function

'Corrige IdCuenta con códigos de cuenta de Comprobantes Tipo
Public Sub UpdateComprobantesTipo()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
   'actualizamos los IdCuenta
   Tbl = " CT_MovComprobante "
   sFrom = " CT_MovComprobante INNER JOIN Cuentas ON CT_MovComprobante.CodCuenta = Cuentas.Codigo "
   sFrom = sFrom & " AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
   sSet = " CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
   sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'no eliminamos los movimientos inconsistentes o los comprobantes vacíos porque si el usuario cambia de plan, podría utilizarlos
   'por esta razón, dejamos la cuenta en cero
   Tbl = " CT_MovComprobante "
   sFrom = " CT_MovComprobante LEFT JOIN Cuentas ON CT_MovComprobante.CodCuenta = Cuentas.Codigo "
   sFrom = sFrom & " AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
   sSet = " CT_MovComprobante.IdCuenta = 0 "
   sWhere = " WHERE Cuentas.IdCuenta IS NULL AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

   
End Sub


Public Function ReadTipoRazFin()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo = 'TIPORAZFIN'"
   Set Rs = OpenRs(DbMain, Q1)
   
   ReDim gTipoRazFin(10)
   i = 0
   
   Do While Rs.EOF = False
   
      If i > UBound(gTipoRazFin) Then
         ReDim Preserve gTipoRazFin(i + 5)
      End If
      
      gTipoRazFin(i).id = vFld(Rs("Codigo"))
      gTipoRazFin(i).Nombre = vFld(Rs("Valor"))
      
      i = i + 1
      Rs.MoveNext
   Loop
      
   Call CloseRs(Rs)

End Function
Public Sub CopyOldIniFile(ByVal FromIniFile As String)
   Dim Buf As String
      
   Buf = GetIniString(FromIniFile, "Config", "ChkVer1", "1")
   Call SetIniString(gIniFile, "Config", "ChkVer1", Buf)

   Buf = GetIniString(FromIniFile, "Config", "ChkVer1", "1")
   Call SetIniString(gIniFile, "Config", "VerNombreCorto", Buf)

   Buf = GetIniString(FromIniFile, "Config", "PathFactura", "")
   Call SetIniString(gIniFile, "Config", "PathFactura", Buf)

   Buf = GetIniString(FromIniFile, "Config", "PathRemu", "")
   Call SetIniString(gIniFile, "Config", "PathRemu", Buf)

   Buf = GetIniString(FromIniFile, "Config", "Printer", "")
   Call SetIniString(gIniFile, "Config", "Printer", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "SelEmprPorRut", "0")
   Call SetIniString(gIniFile, "Opciones", "SelEmprPorRut", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerExento", "1")
   Call SetIniString(gIniFile, "Opciones", "VerExento", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerExento", "1")
   Call SetIniString(gIniFile, "Opciones", "VerExento", Buf)
   
   Buf = GetIniString(FromIniFile, "Opciones", "VerDTE", "1")
   Call SetIniString(gIniFile, "Opciones", "VerDTE", Buf)
   
   Buf = GetIniString(FromIniFile, "Opciones", "VerSucursal", "1")
   Call SetIniString(gIniFile, "Opciones", "VerSucursal", Buf)
   
   Buf = GetIniString(FromIniFile, "Opciones", "VerNumInterno", "1")
   Call SetIniString(gIniFile, "Opciones", "VerNumInterno", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerNumDocHasta", "1")
   Call SetIniString(gIniFile, "Opciones", "VerNumDocHasta", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerMaqReg", "1")
   Call SetIniString(gIniFile, "Opciones", "VerMaqReg", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerCantBoletas", "1")
   Call SetIniString(gIniFile, "Opciones", "VerCantBoletas", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerPropIVA", "1")
   Call SetIniString(gIniFile, "Opciones", "VerPropIVA", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerValCompraHist", "1")
   Call SetIniString(gIniFile, "Opciones", "VerValCompraHist", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerCredArt33", "1")
   Call SetIniString(gIniFile, "Opciones", "VerCredArt33", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerFVenta", "1")
   Call SetIniString(gIniFile, "Opciones", "VerFVenta", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerFUtiliz", "1")
   Call SetIniString(gIniFile, "Opciones", "VerFUtiliz", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerTipoDep", "1")
   Call SetIniString(gIniFile, "Opciones", "VerTipoDep", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerTipoDepHist", "1")
   Call SetIniString(gIniFile, "Opciones", "VerTipoDepHist", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerFechaCompra", "1")
   Call SetIniString(gIniFile, "Opciones", "VerFechaCompra", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerValorInicial", "1")
   Call SetIniString(gIniFile, "Opciones", "VerValorInicial", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerPjeAmortizacion", "1")
   Call SetIniString(gIniFile, "Opciones", "VerPjeAmortizacion", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerFactor", "1")
   Call SetIniString(gIniFile, "Opciones", "VerFactor", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerValorRazonable", "1")
   Call SetIniString(gIniFile, "Opciones", "VerValorRazonable", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerRevalorizacion", "1")
   Call SetIniString(gIniFile, "Opciones", "VerRevalorizacion", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerLCajaDTE", "1")
   Call SetIniString(gIniFile, "Opciones", "VerLCajaDTE", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "VerLCajaNombre", "1")
   Call SetIniString(gIniFile, "Opciones", "VerLCajaNombre", Buf)

   Buf = GetIniString(FromIniFile, "Opciones", "NoDispMsgNewComp", "1")
   Call SetIniString(gIniFile, "Opciones", "NoDispMsgNewComp", Buf)


   Buf = GetIniString(gIniFile, "Cheques", "Ciudad", "")
   Call SetIniString(gIniFile, "Cheques", "Ciudad", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "Banco", "")
   Call SetIniString(gIniFile, "Cheques", "Banco", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "TipoPapel", "1")
   Call SetIniString(gIniFile, "Cheques", "TipoPapel", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "Altura", "4400")
   Call SetIniString(gIniFile, "Cheques", "Altura", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BordeIzq", "2390")
   Call SetIniString(gIniFile, "Cheques", "BordeIzq", Buf)
   
   Buf = GetIniString(gIniFile, "Cheques", "BajarValDig", "250")
   Call SetIniString(gIniFile, "Cheques", "BajarValDig", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BajarFecha", "0")
   Call SetIniString(gIniFile, "Cheques", "BajarFecha", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BajarOrdenDe", "0")
   Call SetIniString(gIniFile, "Cheques", "BajarOrdenDe", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BorrarOrden", "0")
   Call SetIniString(gIniFile, "Cheques", "BorrarOrden", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BajarOrdenDe", "0")
   Call SetIniString(gIniFile, "Cheques", "BajarOrdenDe", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "BorrarPortador", "0")
   Call SetIniString(gIniFile, "Cheques", "BorrarPortador", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "MoverValDig", "0")
   Call SetIniString(gIniFile, "Cheques", "MoverValDig", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "MoverFecha", "0")
   Call SetIniString(gIniFile, "Cheques", "MoverFecha", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "MoverOrdenDe", "0")
   Call SetIniString(gIniFile, "Cheques", "MoverOrdenDe", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "Omitir2DigAno", "0")
   Call SetIniString(gIniFile, "Cheques", "Omitir2DigAno", Buf)

   Buf = GetIniString(gIniFile, "Cheques", "MoverAno", "0")
   Call SetIniString(gIniFile, "Cheques", "MoverAno", Buf)

End Sub

Public Function ResetCompTipoEmpJuntas(Optional ByVal IdEmpresa As Long = 0, Optional ByVal ClearOld As Boolean = False)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Fld As String, Fld2 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
   If IdEmpresa = 0 Then
      IdEmpresa = gEmpresa.id
   End If
   
   If ClearOld Then
      Q1 = " WHERE IdEmpresa = " & IdEmpresa
      Call DeleteSQL(DbMain, "CT_Comprobante", Q1)
      Call DeleteSQL(DbMain, "CT_MovComprobante", Q1)
   End If

   Fld = IdEmpresa & " As IdEmpresa, Correlativo, Nombre, Descrip, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber, IdUsuario, IdComp As IdCompOld "
   Fld2 = " IdEmpresa, Correlativo, Nombre, Descrip, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber, IdUsuario, IdCompOld "
   Q1 = "INSERT INTO CT_Comprobante ( " & Fld2 & " ) SELECT " & Fld & " FROM CT_ComprobanteBase "
   Call ExecSQL(DbMain, Q1)

   Fld = IdEmpresa & " As IdEmpresa, IdComp, Orden, 0 as IdCuenta, CodCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, Conciliado "
   Fld2 = " IdEmpresa, IdComp, Orden, IdCuenta, CodCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, Conciliado "
   Q1 = "INSERT INTO CT_MovComprobante ( " & Fld2 & " ) SELECT " & Fld & " FROM CT_MovComprobanteBase "
   Call ExecSQL(DbMain, Q1)
   
   'reenlazamos los movimientos de comprobantes
   sFrom = " CT_MovComprobante INNER JOIN CT_Comprobante"
   sFrom = sFrom & " ON CT_MovComprobante.IdComp = CT_Comprobante.IdCompOld AND CT_MovComprobante.IdEmpresa = CT_Comprobante.IdEmpresa "
   sSet = " CT_MovComprobante.IdComp = CT_Comprobante.IdComp "
   sWhere = " WHERE CT_MovComprobante.IdEmpresa = " & IdEmpresa
   Call UpdateSQL(DbMain, "CT_MovComprobante", sSet, sFrom, sWhere)
   
   sFrom = " CT_MovComprobante INNER JOIN Cuentas "
   sFrom = sFrom & " ON CT_MovComprobante.CodCuenta = Cuentas.Codigo AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
   sSet = " CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
   sWhere = " WHERE CT_MovComprobante.IdEmpresa = " & IdEmpresa
   sWhere = sWhere & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & gEmpresa.Ano
   Call UpdateSQL(DbMain, "CT_MovComprobante", sSet, sFrom, sWhere)
   
   Call UpdateComprobantesTipo
   
'   'actualizamos los IdCuenta
'   Tbl = " CT_MovComprobante "
'   sFrom = " CT_MovComprobante INNER JOIN Cuentas ON CT_MovComprobante.CodCuenta = Cuentas.Codigo "
'   sFrom = sFrom & " AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
'   sSet = " CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
'   sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
'
'   'no eliminamos los movimientos inconsistentes o los comprobantes vacíos porque si el usuario cambia de plan, podría utilizarlos
'   'por esta razón, dejamos la cuenta en cero
'   Tbl = " CT_MovComprobante "
'   sFrom = " CT_MovComprobante LEFT JOIN Cuentas ON CT_MovComprobante.CodCuenta = Cuentas.Codigo "
'   sFrom = sFrom & " AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
'   sSet = " CT_MovComprobante.IdCuenta = 0 "
'   sWhere = " WHERE Cuentas.IdCuenta IS NULL AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
'   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
'   'reenlazamos los movimientos de comprobantes
'   Q1 = "UPDATE CT_MovComprobante INNER JOIN CT_Comprobante"
'   Q1 = Q1 & " ON CT_MovComprobante.IdComp = CT_Comprobante.IdCompOld AND CT_MovComprobante.IdEmpresa = CT_Comprobante.IdEmpresa "
'   Q1 = Q1 & " SET CT_MovComprobante.IdComp = CT_Comprobante.IdComp "
'   Q1 = Q1 & " WHERE CT_MovComprobante.IdEmpresa = " & IdEmpresa
'   Call ExecSQL(DbMain, Q1)
   
   'actualizamos las cuentas con el nuevo plan, si es que hay
'   Q1 = "UPDATE CT_MovComprobante INNER JOIN Cuentas "
'   Q1 = Q1 & " ON CT_MovComprobante.CodCuenta = Cuentas.Codigo AND CT_MovComprobante.IdEmpresa = Cuentas.IdEmpresa "
'   Q1 = Q1 & " SET CT_MovComprobante.IdCuenta = Cuentas.IdCuenta "
'   Q1 = Q1 & " WHERE CT_MovComprobante.IdEmpresa = " & IdEmpresa
'   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & gEmpresa.Ano
'   Call ExecSQL(DbMain, Q1)
   

End Function


Public Function JoinEmpAno(ByVal DbType As Integer, ByVal Tbl1 As String, ByVal Tbl2 As String, Optional ByVal bAnd As Boolean = True, Optional ByVal SoloEmpresa As Boolean = False) As String
   
   JoinEmpAno = ""
   
   If DbType <> SQL_ACCESS Then

      JoinEmpAno = Tbl1 & ".IdEmpresa = " & Tbl2 & ".IdEmpresa "

      If Not SoloEmpresa Then
         JoinEmpAno = JoinEmpAno & " AND " & Tbl1 & ".Ano = " & Tbl2 & ".Ano"
      End If
      
      If bAnd Then
         JoinEmpAno = " AND " & JoinEmpAno
      End If
      
   End If
   
End Function

Public Function FmtEmprLs(ByVal Rut As Long, ByVal Nombre As String) As String
   Dim sRut As String, l As Integer

   sRut = FmtRut(Rut)
   l = Len(sRut)
      
   l = IIf(l = 11, 13, l + (12 - l) * 1.5)
         
   FmtEmprLs = Right(Space(l) & sRut, l) & vbTab & FCase(Nombre)

End Function

Public Sub HilarPadresPlanCuentasPreDef()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   
'   Q1 = "SELECT Cuentas.IdCuenta, Cuentas.Codigo as CodCuenta, CuentasPadre.IdCuenta, CuentasPadre.Codigo as CodPadre "
'   Q1 = Q1 & " FROM Cuentas  INNER JOIN Cuentas as CuentasPadre ON"
'   Q1 = Q1 & " CuentasPadre.Codigo = concat( left(Cuentas.Codigo, len(Cuentas.Codigo)-2) , '00')"
'   Q1 = Q1 & " AND Cuentas.Idempresa = CuentasPadre.IdEmpresa AND Cuentas.Ano = CuentasPadre.Ano AND CuentasPadre.Nivel = Cuentas.Nivel-1"
'   Q1 = Q1 & " WHERE Cuentas.IdEmpresa = 1 And Cuentas.Ano = 2018 And Cuentas.Nivel = 4"
'   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
'   Set Rs = OpenRs(DbMain, Q1)
   
   'Nivel 4
   Tbl = " Cuentas "
   sFrom = " Cuentas  INNER JOIN Cuentas as CuentasPadre ON "
   sFrom = sFrom & " CuentasPadre.Codigo = " & SqlConcat(gDbType, "Left(Cuentas.Codigo, Len(Cuentas.Codigo)-2)", "'00'")
   sFrom = sFrom & " AND Cuentas.Idempresa = CuentasPadre.IdEmpresa AND Cuentas.Ano = CuentasPadre.Ano AND CuentasPadre.Nivel = Cuentas.Nivel-1"
   sSet = " Cuentas.IdPadre = CuentasPadre.IdCuenta "
   sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano & " AND Cuentas.Nivel = 4"
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
  
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Nivel 3
   Tbl = " Cuentas "
   sFrom = " Cuentas  INNER JOIN Cuentas as CuentasPadre ON "
   sFrom = sFrom & " CuentasPadre.Codigo = " & SqlConcat(gDbType, "left(Cuentas.Codigo, len(Cuentas.Codigo)-4)", "'0000'")
   sFrom = sFrom & " AND Cuentas.Idempresa = CuentasPadre.IdEmpresa AND Cuentas.Ano = CuentasPadre.Ano AND CuentasPadre.Nivel = Cuentas.Nivel-1"
   sSet = " Cuentas.IdPadre = CuentasPadre.IdCuenta "
   sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano & " AND Cuentas.Nivel = 3"
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
  
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   'Nivel 2
   Tbl = " Cuentas "
   sFrom = " Cuentas  INNER JOIN Cuentas as CuentasPadre ON "
   sFrom = sFrom & " CuentasPadre.Codigo = " & SqlConcat(gDbType, "left(Cuentas.Codigo, len(Cuentas.Codigo)-6)", "'000000'")
   sFrom = sFrom & " AND Cuentas.Idempresa = CuentasPadre.IdEmpresa AND Cuentas.Ano = CuentasPadre.Ano AND CuentasPadre.Nivel = Cuentas.Nivel-1"
   sSet = " Cuentas.IdPadre = CuentasPadre.IdCuenta "
   sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano & " AND Cuentas.Nivel = 2"
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
  
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
End Sub

Public Function CreateDir(ByVal Ano As String) As Boolean
   
   CreateDir = False
   
   On Error Resume Next
   Call MkDir(gDbPath & "\Empresas\" & Ano)
   'Err=75 ya existe directorio
   CreateDir = (ERR.Number = 0 Or ERR.Number = 75)
   
End Function

' 6 ene 2020, para reemplazar gImp10 porque cambia retención de boletas de honorarios
' Aumenta en forma gradual de 10.75% hasta 17%
' http://www.sii.cl/destacados/boletas_honorarios/aumento_gradual.html
Public Function ImpBolHono(ByVal Dt As Long, Optional ByVal bComp As Boolean = 0) As Double
   Dim Y As Integer
   
   Y = Year(Dt)
 
   If Y < 2020 Then
      ImpBolHono = 0.1
   ElseIf Y < 2028 Then
      ImpBolHono = 0.1 + 0.0075 * (Y - 2019)
   Else
      ImpBolHono = 0.17
   End If
 
   If bComp Then
      ImpBolHono = 1 - ImpBolHono
   End If
 
End Function



'pipe 2850275
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
      MsgBox1 "Falta especificar correctamente el archivo de configuración de Remuneraciones." & vbCrLf & "Utilice la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
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

'2861570
Public Sub LP_FGr2Clip_Membr(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False)
   
   Call FGr2Clip_membr(Grid, Title, bIncludeCero)

End Sub
'fin 2861570s

'2861570
Public Function Export_SendEmail(Grid As Control, Optional GridTotal0 As Control = Nothing, Optional GridTotal1 As Control = Nothing, Optional GridTotal2 As Control = Nothing, Optional ByVal vnombre As String = "", Optional ByVal Title As String = "", Optional ByVal ColOblig As Integer = -1) As String
   Dim Fd As Long
   Dim SFName As String, fname As String
   Dim Sep As String
   Dim r As Integer
   Dim Buf As String
   Dim FPath As String
   Dim Clip As String
   Dim Membrete As String
     
   'Call FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
   Clip = LP_FGr2String(Grid, Title, False, ColOblig)
   
   If Clip <> "" Then
      '2861570 tema 1
            
      Membrete = "Razón Social " & vbTab & gEmpresa.RazonSocial & vbCrLf
      Membrete = Membrete & " Rut " & vbTab & gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut) & vbCrLf
      Membrete = Membrete & " Dirección " & vbTab & gEmpresa.Direccion & ", " & IIf(gEmpresa.Ciudad <> "", FCase(gEmpresa.Ciudad), FCase(gEmpresa.Comuna)) & vbCrLf
      Membrete = Membrete & " Giro " & vbTab & gEmpresa.Giro & vbCrLf
      Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal1 & vbCrLf
      If gEmpresa.RutRepLegal1 <> "" Then
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & "-" & DV_Rut(gEmpresa.RutRepLegal1) & vbCrLf
      Else
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & vbCrLf
      End If
      
      If gEmpresa.RepConjunta Then
        Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal2 & vbCrLf
        If gEmpresa.RutRepLegal2 <> "" Then
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & "-" & DV_Rut(gEmpresa.RutRepLegal2) & vbCrLf & vbCrLf
        Else
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & vbCrLf & vbCrLf
        End If
      End If
      'fin 2861570 tema 1
         
       Clip = Membrete & Clip
      
      Clip = Clip & FGr2String(GridTotal0)
      Clip = Clip & FGr2String(GridTotal1)
      Clip = Clip & FGr2String(GridTotal2)
   
   On Error Resume Next
      
'   FPath = glHRPathExportPath & "\HRDJ\"
   FPath = gHRPath & "\Email"
   MkDir FPath
'   FPath = gExportPath & "\HRDJ\" & gEmpresa.Rut
   SFName = vnombre & ".xls"
   
   fname = FPath & "\" & SFName

   
   Fd = FreeFile
   ERR.Clear
   
   Open fname For Output As #Fd
   If ERR Then
      MsgErr fname
      Exit Function
   End If

   On Error GoTo 0
      
  Buf = Clip

  Print #Fd, Buf
      
   Close Fd
 

   End If
   Export_SendEmail = fname
End Function
 

