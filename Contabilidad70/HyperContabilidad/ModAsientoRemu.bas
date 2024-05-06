Attribute VB_Name = "ModAsientoRemu"
Option Explicit

Public Const REMU_CTASUELDOBASE = 1 * 10
Public Const REMU_CTAIMPPACT = 1 * 10 + 1          'Imponible pactado Ley 21227
Public Const REMU_CTAHORASEXTRA = 2 * 10
Public Const REMU_CTAHORASDOMINGO = 2 * 10 + 1
Public Const REMU_CTACOMISION = 3 * 10
Public Const REMU_CTABONOSIMP = 4 * 10
Public Const REMU_CTABONOSNOIMP = 5 * 10
Public Const REMU_CTAGRATIFIC = 6 * 10
Public Const REMU_CTASEMANACORRIDA = 6 * 10 + 2    'New: Debe y sumar en remuneraciones por pagar al haber
Public Const REMU_CTAMOVILIZ = 7 * 10
Public Const REMU_CTACOLACION = 8 * 10
Public Const REMU_CTAASIGFAM = 8 * 10 + 2
Public Const REMU_CTACARGASRETRO = 8 * 10 + 3      'Cargas Retroactivas

'2791437
Public Const REMU_CTASSEGCOVIDGASTO = 8 * 10 + 4
'Public Const REMU_CTACOVIDPORPAGAR = 31 * 10
'FIN 2791437

Public Const REMU_CTAAFP = 9 * 10
Public Const REMU_CTAFONASA = 9 * 10 + 2           'New
Public Const REMU_CTAISAPRE = 10 * 10
Public Const REMU_CTAINP = 11 * 10
Public Const REMU_CTASEGACCGASTO = 12 * 10         'seguro accidentes a cuenta de gasto
Public Const REMU_CTASEGACCPPAGAR = 13 * 10        'seguro accidentes a cuenta por pagar
Public Const REMU_CTASEGCESEMPL = 14 * 10          'seguro cesantía por cuenta del trabajador (empleado)
Public Const REMU_CTASEGCESEMPRGASTO = 15 * 10     'seguro cesantía por cuenta del empleador (empresa)  a cuenta de gasto
Public Const REMU_CTASEGCESEMPRPPAGAR = 16 * 10    'seguro cesantía por cuenta del empleador (empresa) a cuenta por pagar
Public Const REMU_CTASISEMPL = 17 * 10             'SIS empleado (menos de 100 empleados)
Public Const REMU_CTASISEMPRGASTO = 18 * 10        'SIS empresa gasto (más de 100 empleados)
Public Const REMU_CTASISEMPRPPAGAR = 19 * 10       'SIS empleado por pagar (más de 100 empleados)
Public Const REMU_CTATRPESADOGASTO = 19 * 10 + 1   'Trabajo Pesado (gasto)
Public Const REMU_CTATRPESADOEMPRESA = 19 * 10 + 2  'Trabajo Pesado Empresa (por pagar)
Public Const REMU_CTATRPESADOTRABAJADOR = 19 * 10 + 3  'Trabajo Pesado Trabajador (por pagar)
Public Const REMU_CTAAPV = 20 * 10
Public Const REMU_CTAAPVCEMPL = 21 * 10            'por cuenta del trabajador
Public Const REMU_CTAAPVCEMPRGASTO = 22 * 10       'por cuenta del empleador a cuenta de gasto
Public Const REMU_CTACOTIZ21227GASTO = 22 * 10 + 1 'cotizaciones por cuenta del empleador por lety 21227
Public Const REMU_CTAAPVCEMPRPPAGAR = 23 * 10      'por cuenta del empleador a cuenta por pagar
'Public Const REMU_CTACCAF_GASTO = 24 * 10         'CCAF cuenta gasto (se elimina por solicitud de Alejandro Contreras 04/jun/12
Public Const REMU_CTACCAF_PORPAGAR = 24 * 10 + 1   'CCAF cuenta por pagar
Public Const REMU_CTACREDCCAF = 25 * 10            'CCAF crédito
Public Const REMU_OTROSDESCUENTOS = 26 * 10
Public Const REMU_CTAVALNOTRABAJADO = 27 * 10      'días no trabajados(permiso, licencia) y minutos de atraso
Public Const REMU_CTAMAYORRETENCION = 28 * 10 - 2  'New: mayor retención (Haber, Cta. Impto. Único)
Public Const REMU_CTAIMPUNICO = 28 * 10
Public Const REMU_CTAANTICIPO = 28 * 10 + 1
Public Const REMU_CTAPRESTAMO = 28 * 10 + 2
Public Const REMU_CTADIFPORCOBRAR = 28 * 10 + 3    'diferencias por cobrar
Public Const REMU_CTARET3PORC = 28 * 10 + 4       ' Retención 3% préstamo solidario
'2791437
'Public Const REMU_CTACOVIDGASTO = 30 * 10
Public Const REMU_CTASSEGCOVIDPORPAGAR = 28 * 10 + 5 'SEG COVID POR PAGAR
'FIN 2791437

'feña 2930020
Public Const REMU_DIFXCOBRAR = 30 * 10
'fin feña 2930020

Public Const REMU_CTAREMPAGAR = 29 * 10


'2791437
'Public Const MAX_CTASREMU = REMU_CTAREMPAGAR

'2930020
'Public Const MAX_CTASREMU = REMU_DIFXCOBRAR
 Public Const MAX_CTASREMU = REMU_CTAREMPAGAR
'2930020
'FIN 2791437
Public gTipoDatosRemu(MAX_CTASREMU) As String

Public Function IniTipoDatosRemu()
 
   gTipoDatosRemu(REMU_CTASUELDOBASE) = "Sueldo Base"
   gTipoDatosRemu(REMU_CTAIMPPACT) = "Imponible Pactado Ley 21227"
   gTipoDatosRemu(REMU_CTAHORASEXTRA) = "Horas Extra"
   gTipoDatosRemu(REMU_CTAHORASDOMINGO) = "Horas Domingo (ordinarias y extra)"
   gTipoDatosRemu(REMU_CTACOMISION) = "Comisión"
   gTipoDatosRemu(REMU_CTABONOSIMP) = "Bonos Imponibles"
   gTipoDatosRemu(REMU_CTABONOSNOIMP) = "Bonos no Imponibles"
   gTipoDatosRemu(REMU_CTAGRATIFIC) = "Gratificación"
   gTipoDatosRemu(REMU_CTASEMANACORRIDA) = "Semana Corrida"
   gTipoDatosRemu(REMU_CTAMOVILIZ) = "Movilización"
   gTipoDatosRemu(REMU_CTACOLACION) = "Colación"
   gTipoDatosRemu(REMU_CTAASIGFAM) = "Asignación Familiar"
   gTipoDatosRemu(REMU_CTACARGASRETRO) = "Cargas Retroactivas"
   
   '2791437
    gTipoDatosRemu(REMU_CTASSEGCOVIDGASTO) = "Seg. Covid (Gasto)"
    'gTipoDatosRemu(REMU_CTACOVIDPORPAGAR) = "Seg. Covid (Por Pagar)"
   'FIN 2791437
   
   gTipoDatosRemu(REMU_CTAAFP) = "AFP"
   gTipoDatosRemu(REMU_CTAFONASA) = "Fonasa"
   gTipoDatosRemu(REMU_CTAISAPRE) = "Isapre"
   gTipoDatosRemu(REMU_CTAINP) = "INP"
   gTipoDatosRemu(REMU_CTASEGACCGASTO) = "Seguro de Accidentes (gasto)"
   gTipoDatosRemu(REMU_CTASEGACCPPAGAR) = "Seguro de Accidentes (por pagar)"
   gTipoDatosRemu(REMU_CTASEGCESEMPL) = "Seguro de Cesantía Trabajador"
   gTipoDatosRemu(REMU_CTASEGCESEMPRGASTO) = "Seguro de Cesantía Empleador (gasto)"
   gTipoDatosRemu(REMU_CTASEGCESEMPRPPAGAR) = "Seguro de Cesantía Empleador (por pagar)"
   gTipoDatosRemu(REMU_CTASISEMPL) = "SIS Trabajador (menos de 100 empleados)"
   gTipoDatosRemu(REMU_CTASISEMPRGASTO) = "SIS Empresa (gasto) (más de 100 empleados)"
   gTipoDatosRemu(REMU_CTASISEMPRPPAGAR) = "SIS Empresa (por pagar) (más de 100 empleados)"
   gTipoDatosRemu(REMU_CTATRPESADOGASTO) = "Trabajo Pesado (gasto)"
   gTipoDatosRemu(REMU_CTATRPESADOEMPRESA) = "Trabajo Pesado Empresa(por pagar)"
   gTipoDatosRemu(REMU_CTATRPESADOTRABAJADOR) = "Trabajo Pesado Trabajador(por pagar)"
   gTipoDatosRemu(REMU_CTAAPV) = "APV"
   gTipoDatosRemu(REMU_CTAAPVCEMPL) = "APVC Trabajador"
   gTipoDatosRemu(REMU_CTAAPVCEMPRGASTO) = "APVC Empleador (gasto)"
   gTipoDatosRemu(REMU_CTACOTIZ21227GASTO) = "Cotizaciones cargo Empleador suspensión Ley 21227"
   gTipoDatosRemu(REMU_CTAAPVCEMPRPPAGAR) = "APVC Empleador (por pagar)"
   'gTipoDatosRemu(REMU_CTACCAF_GASTO) = "CCAF (gasto)"
   gTipoDatosRemu(REMU_CTACCAF_PORPAGAR) = "CCAF (por pagar)"
   gTipoDatosRemu(REMU_CTACREDCCAF) = "Crédito CCAF"
   gTipoDatosRemu(REMU_OTROSDESCUENTOS) = "Otros Descuentos"
   gTipoDatosRemu(REMU_CTAVALNOTRABAJADO) = "Días no Trabajados (licencias, permisos) y min. atraso"
   gTipoDatosRemu(REMU_CTAMAYORRETENCION) = "Mayor Retención"
   gTipoDatosRemu(REMU_CTARET3PORC) = "Retenido 3% Préstamo Solidario"
   gTipoDatosRemu(REMU_CTAIMPUNICO) = "Impuesto Único"
   gTipoDatosRemu(REMU_CTAANTICIPO) = "Anticipo"
   gTipoDatosRemu(REMU_CTAPRESTAMO) = "Prestamo"
   gTipoDatosRemu(REMU_CTADIFPORCOBRAR) = "Diferencias por Cobrar"
   '2791437
    'gTipoDatosRemu(REMU_CTACOVIDGASTO) = "Seg. Covid (Gasto)"
    gTipoDatosRemu(REMU_CTASSEGCOVIDPORPAGAR) = "Seg. Covid (Por Pagar)"
   'FIN 2791437
   
   '2930020
   'gTipoDatosRemu(REMU_DIFXCOBRAR) = "Diferencias por Cobrar"
   'fin 2930020
   
   gTipoDatosRemu(REMU_CTAREMPAGAR) = "Remuneraciones a Pagar"
   
   
   
   
End Function

