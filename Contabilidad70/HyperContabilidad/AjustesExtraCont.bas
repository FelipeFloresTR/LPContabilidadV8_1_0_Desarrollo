Attribute VB_Name = "AjustesExtraCont"
Option Explicit

'tipos de ajustes extra contables
Public Const TAEC_AGREGADOS = 1
Public Const TAEC_DEDUCCIONES = 2
Public Const TAEC_DISPONIBLES = 3
Public Const MAX_TIPOAJUSTESEC = TAEC_DISPONIBLES

'ítems dentro de cada tipo de ajuste
Public Const MAX_ITEMAJUSTESEC = 30
Public Const MAX_CTASAJUSTESEC = 30

Public gTipoAjustesEC(MAX_TIPOAJUSTESEC) As String

Public gAjustesExtraCont(MAX_TIPOAJUSTESEC, MAX_ITEMAJUSTESEC) As ItemAjusteExtraCont_t

'tipos de ingresos de Ajustes Extra Cont
Public Const TIA_CTASASOCIADAS = 1              'Cuentas asociadas: el valor se obtiene sumando los saldos de las cuentas asociadas
Public Const TIA_INGDIRECTO = 2                 'Ingreso Directo
Public Const TIA_CALCULO = 3                    'Se calcula


Type ItemAjusteExtraCont_t
   Nombre As String
   TipoIngresoAjuste As Integer
   IdItemCtasAsociadas As Integer            'Corresponde al Item correlativo de los ajustes que tienen cuentas asociadoas (sólo si TipoIngresoAjuste = TIA_CTASASOCIADAS). Este es el id antiguo de la lista de Ajustes Extracontables que tenían cuentas asociadas. Cuando se agregaron más elementos a la lista, se dejó este ID para no perder las configuraciones anteriores que podía tener el usuario
   AnoDesde As Integer                       'desde qué año es válido
   AnoHasta As Integer                       'hasta que año es válido
End Type

'IdItemCtasAsociadas: número a la izquierda de cada item

'Agregados
'
'1    I.V.A. Pagado en Formulario 29
'2    PPM Pagados
'3    Retiros de Utilidades
'4    Amortización de Capital por pago de cuotas de préstamos recibidos
'5    Gastos Rechazados (Gastos automovil, IDPC, Pago de multas,sueldo de conyuge e hijos, no necesarios para el giro, etc)
'6    Pagos por compras de DS, acciones, cuotas de fondos u otro capital mobiliario
'7    Pagos por compras de activos no depreciables

'     Nuevos desde el 2020
'8    Desarrollo de una actividad agrícola"
'9    Arriendo de bienes raices agrícolas"
'10   Arriendo de bienes raices no agrícolas"
'11   Intereses de depósitos o instrumentos financieros"
'12   Mayor valor en el rescate de cuotas de FM o FI"
'13   Participación en contratos de participación o cuentas de participación"

'
'Deducciones
'
'1    Aporte y Aumentos de Capital
'2    Prestamos Recibidos
'3    Venta de Activos Fijos no depreciables

'     Nuevos desde el 2020
'4    Honorarios Pagados
'5    Impuestos que no sean de la LIR
'6    Gastos afectos al inciso 1°, del art 21 LIR
'7    Gastos afectos al inciso 3°, del art 21 LIR
'8    Créditos incobrables
'9    Pago de IDPC
'10   Pago de IDPC AT 2020 o anteriores que depuran REX
'11   Gastos asociados a INR
'12   Pago 30% ISFUT
'13   Otras partidas pagadas del inciso 2° del art 21, distintos de los anteriores
'14   Ingresos percibidos por la enajenación de bienes depreciables

'
'Cuentas de Disponible
'
'1  Disponible Ingresos / Egresos


Type CtaAjusteExtraCont_t
   IdCuenta(MAX_CTASAJUSTESEC) As Long
   LstCuentas As String
End Type

Public gCtasAjusteExtraCont(MAX_TIPOAJUSTESEC, MAX_ITEMAJUSTESEC) As CtaAjusteExtraCont_t

'tipo de ajuste disponible - iíem Disponibles ingresos/egresos
Public Const TAEC_ITEMDISPONIBLE = 1


Public Sub InitCtasAjustesExtraCont()
   
'   Tipos de ajustes
   gTipoAjustesEC(TAEC_AGREGADOS) = "Agregados"
   gTipoAjustesEC(TAEC_DEDUCCIONES) = "Deducciones"

   If gEmpresa.Ano >= 2020 Then
      gTipoAjustesEC(TAEC_AGREGADOS) = "Ingresos"
      gTipoAjustesEC(TAEC_DEDUCCIONES) = "Egresos"
   End If
   
   gTipoAjustesEC(TAEC_DISPONIBLES) = "Disponibles"

   'AGREGADOS
   
   If gAjustesExtraCont(TAEC_AGREGADOS, 1).Nombre <> "" Then  'ya están seteados todos
      Exit Sub
   End If
   
   gAjustesExtraCont(TAEC_AGREGADOS, 1).Nombre = "IVA Crédito Fiscal Pagado"
   gAjustesExtraCont(TAEC_AGREGADOS, 1).TipoIngresoAjuste = TIA_INGDIRECTO
       
   'sólo hasta 2019
      gAjustesExtraCont(TAEC_AGREGADOS, 2).Nombre = "IVA pagado en F29"
      gAjustesExtraCont(TAEC_AGREGADOS, 2).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 2).IdItemCtasAsociadas = 1
      gAjustesExtraCont(TAEC_AGREGADOS, 2).AnoHasta = 2019
      
      gAjustesExtraCont(TAEC_AGREGADOS, 3).Nombre = "PPM pagado"
      gAjustesExtraCont(TAEC_AGREGADOS, 3).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 3).IdItemCtasAsociadas = 2
      gAjustesExtraCont(TAEC_AGREGADOS, 3).AnoHasta = 2019
      
      gAjustesExtraCont(TAEC_AGREGADOS, 4).Nombre = "Retiro de Utilidades"
      gAjustesExtraCont(TAEC_AGREGADOS, 4).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 4).IdItemCtasAsociadas = 3
      gAjustesExtraCont(TAEC_AGREGADOS, 4).AnoHasta = 2019
      
      gAjustesExtraCont(TAEC_AGREGADOS, 5).Nombre = "Amortización de Capital por pago de cuotas de préstamos recibidos"
      gAjustesExtraCont(TAEC_AGREGADOS, 5).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 5).IdItemCtasAsociadas = 4
      gAjustesExtraCont(TAEC_AGREGADOS, 5).AnoHasta = 2019
         
   gAjustesExtraCont(TAEC_AGREGADOS, 6).Nombre = "Crédito art 33 Bis"
   gAjustesExtraCont(TAEC_AGREGADOS, 6).TipoIngresoAjuste = TIA_CALCULO
      
   'sólo hasta 2019
      gAjustesExtraCont(TAEC_AGREGADOS, 7).Nombre = "Gastos Rechazados (Gastos automovil, IDPC, Pago de multas,sueldo de conyuge e hijos, no necesarios para el giro, etc)"
      gAjustesExtraCont(TAEC_AGREGADOS, 7).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 7).IdItemCtasAsociadas = 5
      gAjustesExtraCont(TAEC_AGREGADOS, 7).AnoHasta = 2019
      
   gAjustesExtraCont(TAEC_AGREGADOS, 8).Nombre = "Incremento por credito de IDPC asociado retiros y dividendos recibidos"
   gAjustesExtraCont(TAEC_AGREGADOS, 8).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 9).Nombre = "Incremento por credito de IDPC asociado retiros y dividendos recibidos sujeto a restitución"
   gAjustesExtraCont(TAEC_AGREGADOS, 9).TipoIngresoAjuste = TIA_INGDIRECTO
      
   'sólo hasta 2019
      gAjustesExtraCont(TAEC_AGREGADOS, 10).Nombre = "Pagos por compras de DS, acciones, cuotas de fondos u otro capital mobiliario"
      gAjustesExtraCont(TAEC_AGREGADOS, 10).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 10).IdItemCtasAsociadas = 6
      gAjustesExtraCont(TAEC_AGREGADOS, 10).AnoHasta = 2019
      
   gAjustesExtraCont(TAEC_AGREGADOS, 11).Nombre = "Ingresos devengados y no percibidos de Empresas Relacionadas al cierre del ejercicio"
   gAjustesExtraCont(TAEC_AGREGADOS, 11).TipoIngresoAjuste = TIA_CALCULO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 12).Nombre = "Ingresos devengados y no percibidos con  plazo mayor a  12 meses desde que  se emitió dcto."
   gAjustesExtraCont(TAEC_AGREGADOS, 12).TipoIngresoAjuste = TIA_CALCULO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 13).Nombre = "Ingresos devengados y no percibidos con plazo mayor a 12 meses desde que pago es  exigible"
   gAjustesExtraCont(TAEC_AGREGADOS, 13).TipoIngresoAjuste = TIA_CALCULO
      
   'sólo hasta 2019
      gAjustesExtraCont(TAEC_AGREGADOS, 14).Nombre = "Pagos por compras de activos no depreciables"
      gAjustesExtraCont(TAEC_AGREGADOS, 14).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 14).IdItemCtasAsociadas = 7
      gAjustesExtraCont(TAEC_AGREGADOS, 14).AnoHasta = 2019
      
   gAjustesExtraCont(TAEC_AGREGADOS, 15).Nombre = "Pago de gastos adeudados antes de ingresar al régimen art 14 ter"
   gAjustesExtraCont(TAEC_AGREGADOS, 15).TipoIngresoAjuste = TIA_CALCULO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 16).Nombre = "Ingreso diferido por incorporación al régimen de la letra A) del art 14 Ter (min 1/5)"
   gAjustesExtraCont(TAEC_AGREGADOS, 16).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 17).Nombre = "Incremento por créditos asociados a Ingreso diferido por incorporación al régimen"
   gAjustesExtraCont(TAEC_AGREGADOS, 17).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 18).Nombre = "Incremento por créditos sujetos a restitución asociados a Ingreso diferido por incorporación al régimen"
   gAjustesExtraCont(TAEC_AGREGADOS, 18).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_AGREGADOS, 19).Nombre = "Otros agregados"
   gAjustesExtraCont(TAEC_AGREGADOS, 19).TipoIngresoAjuste = TIA_INGDIRECTO
    
   gAjustesExtraCont(TAEC_AGREGADOS, 19).Nombre = "Otros agregados"
   gAjustesExtraCont(TAEC_AGREGADOS, 19).TipoIngresoAjuste = TIA_INGDIRECTO
    
   
   'solo desde 2020
      gAjustesExtraCont(TAEC_AGREGADOS, 20).Nombre = "Desarrollo de una actividad agrícola"
      gAjustesExtraCont(TAEC_AGREGADOS, 20).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 20).IdItemCtasAsociadas = 8
      gAjustesExtraCont(TAEC_AGREGADOS, 20).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_AGREGADOS, 21).Nombre = "Arriendo de bienes raices agrícolas"
      gAjustesExtraCont(TAEC_AGREGADOS, 21).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 21).IdItemCtasAsociadas = 9
      gAjustesExtraCont(TAEC_AGREGADOS, 21).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_AGREGADOS, 22).Nombre = "Arriendo de bienes raices no agrícolas"
      gAjustesExtraCont(TAEC_AGREGADOS, 22).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 22).IdItemCtasAsociadas = 10
      gAjustesExtraCont(TAEC_AGREGADOS, 22).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_AGREGADOS, 23).Nombre = "Intereses de depósitos o instrumentos financieros"
      gAjustesExtraCont(TAEC_AGREGADOS, 23).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 23).IdItemCtasAsociadas = 11
      gAjustesExtraCont(TAEC_AGREGADOS, 23).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_AGREGADOS, 24).Nombre = "Mayor valor en el rescate de cuotas de FM o FI"
      gAjustesExtraCont(TAEC_AGREGADOS, 24).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 24).IdItemCtasAsociadas = 12
      gAjustesExtraCont(TAEC_AGREGADOS, 24).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_AGREGADOS, 25).Nombre = "Participación en contratos de participación o cuentas de participación"
      gAjustesExtraCont(TAEC_AGREGADOS, 25).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 25).IdItemCtasAsociadas = 13
      gAjustesExtraCont(TAEC_AGREGADOS, 25).AnoDesde = 2020
         
      gAjustesExtraCont(TAEC_AGREGADOS, 26).Nombre = "Ingresos percibidos por la enajenación de bienes depreciables"
      gAjustesExtraCont(TAEC_AGREGADOS, 26).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_AGREGADOS, 26).IdItemCtasAsociadas = 14
      gAjustesExtraCont(TAEC_AGREGADOS, 26).AnoDesde = 2020
         
         
         
   'DEDUCCIONES
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 1).Nombre = "Saldo Inicial Libro Caja"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 1).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 2).Nombre = "IVA Débito Fiscal Percibido"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 2).TipoIngresoAjuste = TIA_INGDIRECTO
   
   'solo hasta el 2019
      gAjustesExtraCont(TAEC_DEDUCCIONES, 3).Nombre = "Aporte y Aumentos de Capital"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 3).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 3).IdItemCtasAsociadas = 1
      gAjustesExtraCont(TAEC_DEDUCCIONES, 3).AnoHasta = 2019
      
      gAjustesExtraCont(TAEC_DEDUCCIONES, 4).Nombre = "Prestamos Recibidos"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 4).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 4).IdItemCtasAsociadas = 2
      gAjustesExtraCont(TAEC_DEDUCCIONES, 4).AnoHasta = 2019
      
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 5).Nombre = "Gastos presuntos equivalente al 0,5% de los ingresos brutos (Min UTM, Max 15 UTM)"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 5).TipoIngresoAjuste = TIA_CALCULO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 6).Nombre = "Castigo de Deudores Incobrables"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 6).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 7).Nombre = "Pérdida 14 TER  año anterior"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 7).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 8).Nombre = "Costo o inv. actualizada asociado al monto rescatado o enajenado de capital mobiliario, DS, acciones y cuotas de fondos"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 8).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 9).Nombre = "Pago de Ingresos Devengados en años anteriores con Empresas Relacionadas"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 9).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 10).Nombre = "Pago de Ingresos Devengados en años anteriores con plazo mayor a 12 meses desde que se emitió dcto."
   gAjustesExtraCont(TAEC_DEDUCCIONES, 10).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 11).Nombre = "Pago de Ingresos Devengados en años anteriores con plazo mayor a 12 meses desde que pago es exigible."
   gAjustesExtraCont(TAEC_DEDUCCIONES, 11).TipoIngresoAjuste = TIA_INGDIRECTO
   
   'solo hasta 2019
      gAjustesExtraCont(TAEC_DEDUCCIONES, 12).Nombre = "Venta de Activos Fijos no depreciables"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 12).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 12).IdItemCtasAsociadas = 3
      gAjustesExtraCont(TAEC_DEDUCCIONES, 12).AnoHasta = 2019
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 13).Nombre = "Activos Fijos Depreciables a su valor neto, al momento de ingresar al regimen del art. 14 ter"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 13).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 14).Nombre = "Activo Realizable, al momento de ingresar al regimen del art. 14 ter"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 14).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 15).Nombre = "Percepción de ingresos devengados antes de ingresar al regimen art 14 ter"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 15).TipoIngresoAjuste = TIA_CALCULO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 16).Nombre = "Pérdida no absorbida en régimen general al momento de ingresar al régimen art 14 ter"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 16).TipoIngresoAjuste = TIA_INGDIRECTO
   
   gAjustesExtraCont(TAEC_DEDUCCIONES, 17).Nombre = "Otras deducciones"
   gAjustesExtraCont(TAEC_DEDUCCIONES, 17).TipoIngresoAjuste = TIA_INGDIRECTO

   'solo desde el 2020
      
      gAjustesExtraCont(TAEC_DEDUCCIONES, 18).Nombre = "Honorarios Pagados"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 18).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 18).IdItemCtasAsociadas = 4
      gAjustesExtraCont(TAEC_DEDUCCIONES, 18).AnoDesde = 2020
     
      gAjustesExtraCont(TAEC_DEDUCCIONES, 19).Nombre = "Impuestos que no sean de la LIR"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 19).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 19).IdItemCtasAsociadas = 5
      gAjustesExtraCont(TAEC_DEDUCCIONES, 19).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_DEDUCCIONES, 20).Nombre = "Gastos afectos al inciso 1°, del art 21 LIR"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 20).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 20).IdItemCtasAsociadas = 6
      gAjustesExtraCont(TAEC_DEDUCCIONES, 20).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_DEDUCCIONES, 21).Nombre = "Gastos afectos al inciso 3°, del art 21 LIR"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 21).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 21).IdItemCtasAsociadas = 7
      gAjustesExtraCont(TAEC_DEDUCCIONES, 21).AnoDesde = 2020
      
      gAjustesExtraCont(TAEC_DEDUCCIONES, 22).Nombre = "Créditos incobrables"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 22).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 22).IdItemCtasAsociadas = 8
      gAjustesExtraCont(TAEC_DEDUCCIONES, 22).AnoDesde = 2020
   
      gAjustesExtraCont(TAEC_DEDUCCIONES, 23).Nombre = "Pago de IDPC"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 23).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 23).IdItemCtasAsociadas = 9
      gAjustesExtraCont(TAEC_DEDUCCIONES, 23).AnoDesde = 2020
   
      gAjustesExtraCont(TAEC_DEDUCCIONES, 24).Nombre = "Pago de IDPC AT 2020 o anteriores que depuran REX"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 24).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 24).IdItemCtasAsociadas = 10
      gAjustesExtraCont(TAEC_DEDUCCIONES, 24).AnoDesde = 2020
   
      gAjustesExtraCont(TAEC_DEDUCCIONES, 25).Nombre = "Gastos asociados a INR"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 25).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 25).IdItemCtasAsociadas = 11
      gAjustesExtraCont(TAEC_DEDUCCIONES, 25).AnoDesde = 2020
   
      gAjustesExtraCont(TAEC_DEDUCCIONES, 26).Nombre = "Pago 30% ISFUT"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 26).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 26).IdItemCtasAsociadas = 12
      gAjustesExtraCont(TAEC_DEDUCCIONES, 26).AnoDesde = 2020
   
      gAjustesExtraCont(TAEC_DEDUCCIONES, 27).Nombre = "Otras partidas pagadas del inciso 2° del art 21, distintos de los anteriores"
      gAjustesExtraCont(TAEC_DEDUCCIONES, 27).TipoIngresoAjuste = TIA_CTASASOCIADAS
      gAjustesExtraCont(TAEC_DEDUCCIONES, 27).IdItemCtasAsociadas = 13
      gAjustesExtraCont(TAEC_DEDUCCIONES, 27).AnoDesde = 2020
   
   
   
   'Disponibles
   
   gAjustesExtraCont(TAEC_DISPONIBLES, 1).Nombre = "Disponible Ingresos/Egresos"
   gAjustesExtraCont(TAEC_DISPONIBLES, 1).TipoIngresoAjuste = TIA_CTASASOCIADAS
   gAjustesExtraCont(TAEC_DISPONIBLES, 1).IdItemCtasAsociadas = 1

End Sub
Public Sub ReadCtasAjustesExtraCont()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer, j As Integer, k As Integer
   Dim IdItem As Integer
   Dim TipoAjuste As Integer
   Dim LstCuentas As String
      
   For i = 1 To MAX_TIPOAJUSTESEC
      For j = 1 To MAX_ITEMAJUSTESEC
         For k = 1 To MAX_CTASAJUSTESEC
            gCtasAjusteExtraCont(i, j).IdCuenta(k) = 0
         Next k
         gCtasAjusteExtraCont(i, j).LstCuentas = ""
      Next j
   Next i
   
   Q1 = "SELECT IdCtaAjustes, TipoAjuste, IdItem, IdCuenta, CodCuenta FROM CtasAjustesExCont "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY TipoAjuste, IdItem, IdCuenta "
   Set Rs = OpenRs(DbMain, Q1)
     
   i = 1
   LstCuentas = ""
   IdItem = 0
   TipoAjuste = 0
   
   Do While Not Rs.EOF
   
      If TipoAjuste <> vFld(Rs("TipoAjuste")) Then
         TipoAjuste = vFld(Rs("TipoAjuste"))
         IdItem = vFld(Rs("IdItem"))
         i = 1
         LstCuentas = ""
         
      ElseIf IdItem <> vFld(Rs("IdItem")) Then
         IdItem = vFld(Rs("IdItem"))
         i = 1
         LstCuentas = ""
      End If
      
      LstCuentas = LstCuentas & "," & vFld(Rs("IdCuenta"))
      gCtasAjusteExtraCont(vFld(Rs("TipoAjuste")), vFld(Rs("IdItem"))).IdCuenta(i) = vFld(Rs("IdCuenta"))
      gCtasAjusteExtraCont(vFld(Rs("TipoAjuste")), vFld(Rs("IdItem"))).LstCuentas = LstCuentas & ","       ',1,2,3,"
      
      i = i + 1
      
      If i > 30 Then
         MsgBox1 "Se ha superado la cantidad de cuentas para el Item " & gAjustesExtraCont(vFld(Rs("TipoAjuste")), vFld(Rs("IdItem"))).Nombre & " de Ajustes Extracontables.", vbExclamation
         Exit Do
      End If
         
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
End Sub
Public Function LoadValCuentasAjustes(ByVal TipoAjusteExtraCont As Integer, ByVal IdItemCuentas As Integer, Optional ByVal TipoComp As Integer = 0) As Double
   Dim LstCuentas As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Tot As Double
   
   LstCuentas = gCtasAjusteExtraCont(TipoAjusteExtraCont, IdItemCuentas).LstCuentas
   
   If LstCuentas = "" Then
      LoadValCuentasAjustes = 0
      Exit Function
   End If
   LstCuentas = Mid(LstCuentas, 2)
   LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
   
   If TipoAjusteExtraCont = TAEC_AGREGADOS Then
      Q1 = "SELECT Sum(Debe) "
   Else
      Q1 = "SELECT Sum(Haber) "
   End If
   Q1 = Q1 & " as SumCtas FROM MovComprobante INNER JOIN Comprobante ON (MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " WHERE IdCuenta IN (" & LstCuentas & ")"
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Tipo <> " & TC_APERTURA            'se agrega esta condición por solicitud de Joshua Nicolás Catrin (07/09/2018)
   If TipoComp > 0 Then
      Q1 = Q1 & " AND Comprobante.Tipo = " & TipoComp
   End If
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Tot = 0
   If Not Rs.EOF Then
      Tot = vFld(Rs("SumCtas"))
   End If
   
   Call CloseRs(Rs)
   
   LoadValCuentasAjustes = Tot
   
End Function

Public Function LoadValCuentasAjustes14D(ByVal TipoAjusteExtraCont As Integer, ByVal IdItemCuentas As Integer, Optional ByVal TipoComp As Integer = 0, Optional ByVal LstCuentasIn As String, Optional ByVal idcomp As Long = 0) As Double
   Dim LstCuentas As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Tot As Double
   
   If IdItemCuentas <> 0 Then
      LstCuentas = gCtasAjusteExtraCont(TipoAjusteExtraCont, IdItemCuentas).LstCuentas
      
      If LstCuentas = "" Then
         LoadValCuentasAjustes14D = 0
         Exit Function
      End If
      LstCuentas = Mid(LstCuentas, 2)
      LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
      
   Else
   '3076934
' ANTES
'      LstCuentas = LstCuentasIn
'      LstCuentas = Mid(LstCuentas, 2)
'
'      If LstCuentas = "" Then
'         LoadValCuentasAjustes14D = 0
'         Exit Function
'      End If
' AHORA
     LstCuentas = LstCuentasIn
     
     If Left(LstCuentas, 1) = "," Then
       LstCuentas = Mid(LstCuentas, 2)
     End If

      If Mid(LstCuentas, 2) = "" Then
        LoadValCuentasAjustes14D = 0
        Exit Function
      End If
'FIN
   End If
   
   
   If TipoAjusteExtraCont = TAEC_AGREGADOS Then   'Ingresos
      Q1 = "SELECT Sum(Haber) "
      
   Else           'DEDUCCIONES
      Q1 = "SELECT Sum(Debe) "
      
      If IdItemCuentas = 8 Then         ' 8: Créditos incobrables
         TipoComp = TC_TRASPASO                                                                    'en duro, no queda otra
      End If
   End If
   
   Q1 = Q1 & " as SumCtas FROM MovComprobante INNER JOIN Comprobante ON (MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " WHERE IdCuenta IN (" & LstCuentas & ")"
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Tipo <> " & TC_APERTURA            'se agrega esta condición por solicitud de Joshua Nicolás Catrin (07/09/2018)
   
   If TipoComp > 0 Then
      Q1 = Q1 & " AND Comprobante.Tipo = " & TipoComp
   End If
   
   If idcomp > 0 Then
      Q1 = Q1 & " AND Comprobante.IdComp = " & idcomp
   End If
   
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Tot = 0
   If Not Rs.EOF Then
      Tot = vFld(Rs("SumCtas"))
   End If
   
   Call CloseRs(Rs)
   
   LoadValCuentasAjustes14D = Tot
   
End Function

Public Function GetVal33Bis() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Val33bis As Double
   Dim ValUTM As Double
   Dim Val500UTM As Double
   
   GetVal33Bis = 0
   
   ValUTM = GetValUTM()
   
   Val500UTM = 500 * ValUTM

   
   Q1 = "SELECT CredArt33Bis FROM EmpresasAno "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Val33bis = vFld(Rs("CredArt33Bis"))
   End If
   
   Call CloseRs(Rs)
   
   If Val33bis > Val500UTM Then
      GetVal33Bis = Val500UTM
   Else
      GetVal33Bis = Val33bis
   End If
      

End Function

Public Function GetValUTM() As Double
   Dim ValUTM As Double
   Dim Fecha As Long

   GetValUTM = 0
   Fecha = DateSerial(gEmpresa.Ano, 12, 31)
      
   If GetValMoneda("UTM", ValUTM, Fecha, False) = True Then     'obtiene la última UTM ingresada en el sistema que sea a lo más del 31 dic del año actual}
      GetValUTM = ValUTM
   Else
      MsgBox1 "No se encontró el valor de la UTM para calcular los topes de Crádito 33 Bis y Gastos Presuntos.", vbExclamation
   End If

End Function


'pipe 2699582

Public Function GetValPpm() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ValPpm As Double
   Dim FechaPPM As Long
   Dim TipoPPM As Boolean
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='PPM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TipoPPM = IIf(vFld(Rs("Valor")) <> 0, True, False)
      FechaPPM = DateSerial(gEmpresa.Ano, 1, 20)
   End If
   
   Call CloseRs(Rs)
   
   GetValPpm = 0
   
   If TipoPPM Then
    Q1 = "SELECT sum(IIF(COMPROBANTE.FECHA <= " & FechaPPM & " AND ParamEmpresa.tipo = 'CTAPPMOBLI', 0, MovComprobante.DEBE)) AS VALOR "
   Else
    Q1 = "SELECT sum(MovComprobante.DEBE) AS VALOR "
   End If
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo in ('CTAPPMOBLI','CTAPPMVOLU') "
   If gDbType = SQL_SERVER Then
   Q1 = Q1 & " AND  MovComprobante.idCuenta = ParamEmpresa.valor and Comprobante.Ano = " & gEmpresa.Ano
   Else
   Q1 = Q1 & " AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) and Comprobante.Ano = " & gEmpresa.Ano
   End If
   
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO

     
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetValPpm = vFld(Rs("VALOR"))
   End If
   
    GetValPpm = (LoadAllPpmObli + LoadAllPpmVolun) - GetValPpm
   
     
   
   Call CloseRs(Rs)
   
End Function

Public Function LoadAllPpmObli() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim FechaPPM As Long
   Dim TipoPPM As Boolean
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='PPM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TipoPPM = IIf(vFld(Rs("Valor")) <> 0, True, False)
      FechaPPM = DateSerial(gEmpresa.Ano, 1, 20)
   End If
   
   Call CloseRs(Rs)
   
   Dim montoActualizado As String
     
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMOBLI' "
   
    If gDbType = SQL_SERVER Then
   Q1 = Q1 & " AND  MovComprobante.idCuenta = ParamEmpresa.valor"
   Else
   Q1 = Q1 & " AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) "
   End If
   
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
   
   If TipoPPM Then
        Q1 = Q1 & " AND COMPROBANTE.FECHA > " & FechaPPM
   End If

   Set Rs = OpenRs(DbMain, Q1)

   Do While Rs.EOF = False
    
           Q2 = "SELECT Factor "
           Q2 = Q2 & " FROM FactorActAnual"
           Q2 = Q2 & " WHERE Ano = " & Year(vFld(Rs("Fecha")))
           Q2 = Q2 & " AND MesCol = 12 "
           Q2 = Q2 & " AND MesRow = " & month(vFld(Rs("Fecha")))
        
           Set Rs2 = OpenRs(DbMain, Q2)
          If Rs2.EOF = False Then
              
              montoActualizado = Format(vFld(Rs("DEBE")) * vFld(Rs2("Factor")), NUMFMT)
          Else
             
             montoActualizado = Format(vFld(Rs("DEBE")) * 1, NUMFMT)
             
          End If
          Call CloseRs(Rs2)
          LoadAllPpmObli = LoadAllPpmObli + montoActualizado
          
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
    
     

End Function

Public Function LoadAllPpmVolun() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim montoActualizado As String
  
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMVOLU' "
   
    If gDbType = SQL_SERVER Then
   Q1 = Q1 & " AND  MovComprobante.idCuenta = ParamEmpresa.valor"
   Else
   Q1 = Q1 & " AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) "
   End If
   
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO

   Set Rs = OpenRs(DbMain, Q1)
   
  
   Do While Rs.EOF = False

      
      Q2 = "SELECT Factor "
           Q2 = Q2 & " FROM FactorActAnual"
           Q2 = Q2 & " WHERE Ano = " & Year(vFld(Rs("Fecha")))
           Q2 = Q2 & " AND MesCol = 12 "
           Q2 = Q2 & " AND MesRow = " & month(vFld(Rs("Fecha")))
        
           Set Rs2 = OpenRs(DbMain, Q2)
          If Rs2.EOF = False Then
              
            montoActualizado = Format(vFld(Rs("DEBE")) * vFld(Rs2("Factor")), NUMFMT)
          Else
            montoActualizado = Format(vFld(Rs("DEBE")) * 1, NUMFMT)
          End If
          Call CloseRs(Rs2)
          
          LoadAllPpmVolun = LoadAllPpmVolun + montoActualizado

      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
     

End Function

'fin 2699582

