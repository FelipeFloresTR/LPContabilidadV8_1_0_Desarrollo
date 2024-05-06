Attribute VB_Name = "AjustesExtraContRLI_RAB"
Option Explicit

'tipos de ajustes extra contables RLI-RAB
Public Const TAECR_AGREGADOSPAGADOS = 1
Public Const TAECR_AGREGADOSADEUDADOS = 2
Public Const TAECR_DEDUCCIONESPERCIBIDAS = 3
Public Const TAECR_DEDUCCIONESDEVENGADAS = 4
Public Const MAX_TIPOAJUSTESECRLI = TAECR_DEDUCCIONESDEVENGADAS

'ítems dentro de cada tipo de ajuste
Public Const MAX_GRUPOAJUSTESECRLI = 6
Public Const MAX_ITEMAJUSTESECRLI = 20
Public Const MAX_CTASAJUSTESECRLI = 30

Public gTipoAjustesECRLI(MAX_TIPOAJUSTESECRLI) As String
Public gGrupoAjustesECRLI(MAX_TIPOAJUSTESECRLI, MAX_GRUPOAJUSTESECRLI) As String

Public gAjustesExtraContRLI(MAX_TIPOAJUSTESECRLI, MAX_GRUPOAJUSTESECRLI, MAX_ITEMAJUSTESECRLI) As AjusteExtraContRLI_t

''tipos de ingresos de Ajustes Extra Cont
'Public Const TIA_CTASASOCIADAS = 1              'Cuentas asociadas: el valor se obtiene sumando los saldos de las cuentas asociadas
'Public Const TIA_INGDIRECTO = 2                 'Ingreso Directo
'Public Const TIA_CALCULO = 3                    'Se calcula


Type AjusteExtraContRLI_t
   Nombre As String
   TipoItem As String
   orden As Integer
   IdCuenta(MAX_CTASAJUSTESECRLI) As Long
   LstCuentas As String
End Type




'IdItemCtasAsociadas: número a la izquierda de cada item

'Agregados
'
'1  I.V.A. Pagado en Formulario 29
'2  PPM Pagados
'3  Retiros de Utilidades
'4  Amortización de Capital por pago de cuotas de préstamos recibidos
'5  Gastos Rechazados (Gastos automovil, IDPC, Pago de multas,sueldo de conyuge e hijos, no necesarios para el giro, etc)
'6  Pagos por compras de DS, acciones, cuotas de fondos u otro capital mobiliario
'7  Pagos por compras de activos no depreciables
'
'Deducciones
'
'1  Aporte y Aumentos de Capital
'2  Prestamos Recibidos
'3  Venta de Activos Fijos no depreciables
'
'Cuentas de Disponible
'
'1  Disponible Ingresos / Egresos


'Type CtaAjusteExtraCont_t
'   IdCuenta(MAX_CTASAJUSTESEC) As Long
'   LstCuentas As String
'End Type



Public Sub InitCtasAjustesExtraContRLI()
   Dim i As Integer
   
'   Tipos de ajustes
   gTipoAjustesECRLI(TAECR_AGREGADOSPAGADOS) = "Agregados Pagados"
   gTipoAjustesECRLI(TAECR_AGREGADOSADEUDADOS) = "Agregados Adeudados"
   gTipoAjustesECRLI(TAECR_DEDUCCIONESPERCIBIDAS) = "Deducciones Percibidas"
   gTipoAjustesECRLI(TAECR_DEDUCCIONESDEVENGADAS) = "Deducciones Devengadas"

'  Grupos
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 1) = "Gastos Rechazados afectos a IU Art. 21 o IGC/Adicional"
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 2) = "Gastos Rechazados no Afecto (NA)"
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 3) = "Gastos Rechazados solo aumentan (SIN EFECTO)"
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 4) = "Gastos Imputables a Ingresos no Renta (REX-INR)"
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 5) = "Gastos Imputables  a Ingresos Exentos (REX-INR)"
   gGrupoAjustesECRLI(TAECR_AGREGADOSPAGADOS, 6) = "Donaciones"
   
   gGrupoAjustesECRLI(TAECR_AGREGADOSADEUDADOS, 1) = "Impuestos"
   gGrupoAjustesECRLI(TAECR_AGREGADOSADEUDADOS, 2) = "Provisiones Varias"
   gGrupoAjustesECRLI(TAECR_AGREGADOSADEUDADOS, 3) = "Estimaciones varias y ajustes contables"
   gGrupoAjustesECRLI(TAECR_AGREGADOSADEUDADOS, 4) = "Depreciación"

   gGrupoAjustesECRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1) = "Otras Deducciones Permitidas"
   gGrupoAjustesECRLI(TAECR_DEDUCCIONESPERCIBIDAS, 2) = "REX-INR (Propios)"
   gGrupoAjustesECRLI(TAECR_DEDUCCIONESPERCIBIDAS, 3) = "REX-EX (Propios)"

   gGrupoAjustesECRLI(TAECR_DEDUCCIONESDEVENGADAS, 1) = "Deducciones Varias"
   gGrupoAjustesECRLI(TAECR_DEDUCCIONESDEVENGADAS, 2) = "Rentas en Zonas Extremas - Regimenes Preferenciales"

   '------------------------------------------------------------------------------------------------------------
   'Agregados Pagados
   
   'Grupo 1: "Gastos Rechazados afectos a IU Art. 21 o IGC/Adicional"
   
   For i = 1 To 7
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, i).orden = i
   Next i
   

   If gEmpresa.Ano < 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 1).Nombre = "Pérdidas Sufrida por el Negocio"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 2).Nombre = "Sueldos del Conyuge, hijos solteros menoras de 18 años y excesos…."
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 3).Nombre = "Indemnización Años de Servicios"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 4).Nombre = "Otras Remuneraciones no aceptadas"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 5).Nombre = "Gastos de Automovil"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 6).Nombre = "Reajustes, Intereses y Diferencia de Cambio"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 7).Nombre = "Cualquier otro gasto no aceptado afecto al articulo 21"
   Else
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 6).Nombre = "Gastos que beneficien a relacionados o propietarios"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, 7).Nombre = "Gastos sin acreditar naturaleza y efectividad o rechazados por norma especial"
   End If
   
   
   For i = 1 To 7
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 1, i).TipoItem = "A" & i
   Next i
   
   'Grupo 2: "Gastos Rechazados no Afecto (NA)"
   For i = 1 To 9
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 1).Nombre = "Intereses, Reajustes y Multas Fiscales"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 2).Nombre = "Impuesto de Primera Categoria"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 3).Nombre = "Impuesto Territorial"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 4).Nombre = "Impuesto de 1ra. Categoria Voluntario Pagado"
   If gEmpresa.Ano < 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 5).Nombre = "Otros Impuestos"
   End If
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 6).Nombre = "Impuesto Unico 40%"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 7).Nombre = "Patentes Mineras"
   
   If gEmpresa.Ano >= 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 8).Nombre = "Partidas del inc. 1° no afectas al IU , del art 21 LIR"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 8).orden = 1
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, 9).Nombre = "IDPC AT 2020 o anterior que depure REX"
      For i = 1 To 7
         gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, i).orden = i + 1
      Next i
      
   End If
   
   For i = 1 To 9
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 2, i).TipoItem = "B" & i
   Next i
   
   'Grupo 3: "Gastos Rechazados solo aumentan (SIN EFECTO)"
   
   For i = 1 To 2
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 3, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 3, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 3, 1).Nombre = "Sumas que deben imputarse al costo de los bienes"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 3, 2).Nombre = "Gtos. Anticipados a imputar a ejercicios siguientes"
   
   For i = 1 To 2
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 3, i).TipoItem = "C" & i
   Next i
   
   'Grupo 4: "Gastos Imputables a Ingresos no Renta (REX-INR)"
   For i = 1 To 4
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, 1).Nombre = "Gtos. Imputables a ingresos NO Renta"
   
   If gEmpresa.Ano < 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, 2).Nombre = "Impuesto Unico por Utilidades Afectas Recibidas"
   Else
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, 2).Nombre = "Proporcionalidad gastos INR"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, 3).Nombre = "ISFUT pagado durante 2020"
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, 4).Nombre = "Gastos financieros y otros asociados"
   End If
  
   For i = 1 To 4
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 4, i).TipoItem = "D" & i
   Next i
   
   'Grupo 5: "Gastos Imputables  a Ingresos Exentos (REX-INR)"
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 5, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 5, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 5, 1).Nombre = "Gtos. Imputables a ingresos Exentos"
   
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 5, i).TipoItem = "E" & i
   Next i
   
   
   'Grupo 5: "Donaciones"
   
   '3402617
    If gEmpresa.R14ASemiIntegrado = True And gEmpresa.Ano >= 2023 Then
       For i = 1 To 17
           gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).Nombre = ""
           gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).orden = i
        Next i
    Else
        For i = 1 To 16
           gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).Nombre = ""
           gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).orden = i
        Next i
  
    End If
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 1).Nombre = "Fines Culturales Art. 8 Ley 18.985/90"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 2).Nombre = "Fines Educacionales Art. 3 Ley 19.247/93"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 3).Nombre = "Fines Deportivos Art. 62  Ley 19.712, Cred 50%"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 4).Nombre = "Fines Deportivos Art. 62  Ley 19.712, Cred 30%"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 5).Nombre = "Fines Sociales Art. 1° Ley 19.885/03"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 6).Nombre = "Fines Sociales para prevenir adicciones de alcohol y drogas"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 7).Nombre = "Universidades e IP Art. 69 Ley 18.681/87"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 8).Nombre = "Art. 46 DL 3.063"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 9).Nombre = "DL 45/73"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 10).Nombre = "Fundacion Teresa de los Anes"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 11).Nombre = "Art. 31 N° 7 LIR"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 12).Nombre = "CORFO Ley 6.640/41"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 13).Nombre = "FNR Según Art. 4 Ley 20.444/10"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 14).Nombre = "Catastrofes art. 7° Ley 16.282/85"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 15).Nombre = "Según Ley 21.015"
   gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 16).Nombre = "Art. 37 DL 1.939"
   
   '3402617
    If gEmpresa.R14ASemiIntegrado = True And gEmpresa.Ano >= 2023 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, 17).Nombre = "Donaciones Bs Inmubles Plan Emerg. Habitacional Ley 21.450"

      For i = 1 To 17
        gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).TipoItem = "F" & i
      Next i

     Else
        For i = 1 To 16
           gAjustesExtraContRLI(TAECR_AGREGADOSPAGADOS, 6, i).TipoItem = "F" & i
        Next i
    End If
    '3402617
 
 
   '------------------------------------------------------------------------------------------------------------
   'Agregados Adeudados
   
   'Grupo 1: "Impuestos"
   For i = 1 To 6
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, i).orden = i
   Next i
      
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 1).Nombre = "Impuesto de Primera Categoría"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 2).Nombre = "Impuesto Territorial"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 3).Nombre = "Otros Impuestos"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 4).Nombre = "Impuesto Unico 40%"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 5).Nombre = "Patentes Mineras"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, 6).Nombre = "Provision Impuesto de 1ra. Categoría Voluntario"
   
   For i = 1 To 6
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 1, i).TipoItem = "G" & i
   Next i
      
   'Grupo 2: "Provisiones Varias"
   
   For i = 1 To 12
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, i).orden = i
   Next i

   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 1).Nombre = "Intereses, reajustes y Multas fiscales"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 2).Nombre = "Reajustes. Intereses y Diferencia de cambio"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 3).Nombre = "Perdidas Sufridas por el Negocio"
   If gEmpresa.Ano < 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 4).Nombre = "Sueldos del Cónyuge, hijos solteros menores de 18 años y exceso sueldo empresarial"
   Else
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 4).Nombre = ""
   End If
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 5).Nombre = "Indemnización años de servicios"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 6).Nombre = "Otras Remuneraciones no aceptadas"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 7).Nombre = "Gastos de automóvil"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 8).Nombre = "Otros no autorizados por Art. 31 que importan flujos futuros"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 9).Nombre = "Intereses Rechazados como Gasto"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 10).Nombre = "Gratificaciones y Participaciones Voluntarias"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 11).Nombre = "Provisión Vacaciones"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, 12).Nombre = "Otras Provisiones"
   
   For i = 1 To 12
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 2, i).TipoItem = "H" & i
   Next i
   
   
   'Grupo 3: "Estimaciones varias y ajustes contables"
   For i = 1 To 14
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, i).orden = i
   Next i
   
   If gEmpresa.Ano < 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 1).Nombre = "Estimación deuda incobrable"
   Else
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 1).Nombre = "Estimación  y/o castigos de deudas incobrables, según criterios financieros"
   End If
   
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 2).Nombre = "Castigo deuda no aceptada"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 3).Nombre = "Condonación deudas no aceptadas"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 4).Nombre = "Otras Depreciaciones No Aceptadas"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 5).Nombre = "Corrección Monetaria Financiera"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 6).Nombre = "Ajustes por Contratos Leasing"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 7).Nombre = "Otras diferencias temporales financieras que no importan flujo"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 8).Nombre = "Rentas Tributables no reconocidas financieramente"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 9).Nombre = "Correccion Monetaria, disminuciones del Capital Propio"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 10).Nombre = "Correccion Monetaria, activos no monetarios"
   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 11).Nombre = "Otros Ajustes"
   If gEmpresa.Ano >= 2020 Then
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 12).Nombre = "Ingreso diferido por cambio de régimen"
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 13).Nombre = "Intereses devengados en bonos del art. 104 LIR"
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 14).Nombre = "Ingresos devengados por cambio de régimen"
      For i = 12 To 14
         gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, i).orden = i - 1
      Next i
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, 11).orden = 14
   End If
   
   For i = 1 To 14
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 3, i).TipoItem = "I" & i
   Next i
   
   'Grupo 4: "Depreciación"
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 4, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 4, i).orden = i
   Next i

   gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 4, 1).Nombre = "Depreciación Financiera"
   
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_AGREGADOSADEUDADOS, 4, i).TipoItem = "J" & i
   Next i
   
   
   '------------------------------------------------------------------------------------------------------------
   'Deducciones Percibidas
   
   'Grupo 1: "Otras Deducciones Permitidas"
   For i = 1 To 3
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, 1).Nombre = "Utilidad reconocida, Renta Presunta (Propia)"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, 2).Nombre = "Otras Deducciones Permitidas"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, 3).Nombre = "Rentas Exentas 1° Categoría afectas a global (Propia)"
   
   For i = 1 To 3
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 1, i).TipoItem = "K" & i
   Next i
      
   'Grupo 2: "REX-INR (Propios)"
   
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 2, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 2, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 2, 1).Nombre = "Ingreso no constitutivos de renta (art. 17)"
  
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 2, i).TipoItem = "L" & i
   Next i
  
   'Grupo 3: "REX-EX (Propios)"
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 3, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 3, i).orden = i
   Next i
   
   gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 3, 1).Nombre = "Rentas exentas 1 a. cat y exentas a global"
   
   For i = 1 To 1
      gAjustesExtraContRLI(TAECR_DEDUCCIONESPERCIBIDAS, 3, i).TipoItem = "M" & i
   Next i
         
   '------------------------------------------------------------------------------------------------------------
   'Deducciones Devengadas
   
   'Grupo 1: "Deducciones Varias"
   
   For i = 1 To 13
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, i).orden = i
   Next i

   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 1).Nombre = "Reconocimiento de utilidad, renta presunta"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 2).Nombre = "Indeminización años de servicios"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 3).Nombre = "Ajustes por Contratos Leasing"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 4).Nombre = "Otras Diferencias Temporales"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 5).Nombre = "Otras Deducciones Permitidas"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 6).Nombre = "Gasto Goodwill tributario del Ejercicio"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 7).Nombre = "Impuesto Especifico a la Actividad Minera"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 8).Nombre = "Correccion Moneteria, Capital Propio Inicial"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 9).Nombre = "Correccion Monetaria, Aumentos de Capital Propio"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 10).Nombre = "Correccion Monetaria, Pasivos no Monetarios"
   
   If gEmpresa.Ano >= 2020 Then
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 11).Nombre = "Gastos adeudados por cambio de régimen"
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 11).orden = 1
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 12).Nombre = "Castigo de deudas incobrables, art. 31 N° 4 LIR"
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 12).orden = 2
      
      '3402617
      If gEmpresa.R14ASemiIntegrado = True And gEmpresa.Ano >= 2023 Then

      Else
        gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 13).Nombre = "Amortización de intangibles, art. 22° transitorio bis, ley 21.210"
        gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 13).orden = 3
      End If
      
'      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 13).Nombre = "Amortización de intangibles, art. 22° transitorio bis, ley 21.210"
'      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, 13).orden = 3
      '3402617
      
      For i = 1 To 10
         gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, i).orden = i + 3
      Next i
   End If
   
   '3402617
      If gEmpresa.R14ASemiIntegrado = True And gEmpresa.Ano >= 2023 Then
          For i = 1 To 12
           gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, i).TipoItem = "N" & i
          Next i

      Else
         For i = 1 To 13
           gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 1, i).TipoItem = "N" & i
         Next i
      End If
   '3402617
   
   'Grupo 2: "Rentas en Zonas Extremas - Regimenes Preferenciales"
   For i = 1 To 4
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, i).Nombre = ""
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, i).orden = i
   Next i

   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, 1).Nombre = "Rentas Exentas Ley 18.392/1985 - Ley Navarino"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, 2).Nombre = "Rentas Exentas Ley 19.149/1992 - Ley Tierra del Fuego"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, 3).Nombre = "Rentas Exentas Ley 19.709/2001 - Ley de Tocopilla"
   gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, 4).Nombre = "Otras Rentas Exentas de IDPC, Afectas a IGC o IA"
   
   For i = 1 To 4
      gAjustesExtraContRLI(TAECR_DEDUCCIONESDEVENGADAS, 2, i).TipoItem = "P" & i
   Next i

End Sub

Public Sub ReadCtasAjustesExtraContRLI()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer, j As Integer, k As Integer, l As Integer
   Dim IdGrupo As Integer, IdItem As Integer
   Dim TipoAjuste As Integer
   Dim LstCuentas As String
      
   For i = 1 To MAX_TIPOAJUSTESECRLI
      For j = 1 To MAX_GRUPOAJUSTESECRLI
         For k = 1 To MAX_ITEMAJUSTESECRLI
            For l = 1 To MAX_CTASAJUSTESECRLI
               gAjustesExtraContRLI(i, j, k).IdCuenta(l) = 0
            Next l
            gAjustesExtraContRLI(i, j, k).LstCuentas = ""
         Next k
      Next j
   Next i
   
   Q1 = "SELECT IdCtaAjustesrli, TipoAjuste, IdGrupo, IdItem, IdCuenta, CodCuenta FROM CtasAjustesExContRLI "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY TipoAjuste, IdGrupo, IdItem, IdCuenta "
   Set Rs = OpenRs(DbMain, Q1)
     
   i = 1
   LstCuentas = ""
   IdGrupo = 0
   IdItem = 0
   TipoAjuste = 0
   
   Do While Not Rs.EOF
   
      If TipoAjuste <> vFld(Rs("TipoAjuste")) Then
         TipoAjuste = vFld(Rs("TipoAjuste"))
         IdGrupo = vFld(Rs("IdGrupo"))
         IdItem = vFld(Rs("IdItem"))
         i = 1
         LstCuentas = ""
         
      ElseIf IdGrupo <> vFld(Rs("IdGrupo")) Then
         IdGrupo = vFld(Rs("IdGrupo"))
         IdItem = vFld(Rs("IdItem"))
         i = 1
         LstCuentas = ""
      
      ElseIf IdItem <> vFld(Rs("IdItem")) Then
         IdItem = vFld(Rs("IdItem"))
         i = 1
         LstCuentas = ""
      End If
      
      LstCuentas = LstCuentas & "," & vFld(Rs("IdCuenta"))
      gAjustesExtraContRLI(vFld(Rs("TipoAjuste")), vFld(Rs("IdGrupo")), vFld(Rs("IdItem"))).IdCuenta(i) = vFld(Rs("IdCuenta"))
      gAjustesExtraContRLI(vFld(Rs("TipoAjuste")), vFld(Rs("IdGrupo")), vFld(Rs("IdItem"))).LstCuentas = LstCuentas & ","       '",1,2,3,"
      
      i = i + 1
      
      If i > 30 Then
         MsgBox1 "Se ha superado la cantidad de cuentas para el Item " & gAjustesExtraContRLI(vFld(Rs("TipoAjuste")), vFld(Rs("IdGrupo")), vFld(Rs("IdItem"))).Nombre & " de Ajustes Extracontables.", vbExclamation
         Exit Do
      End If
         
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
End Sub

