Attribute VB_Name = "BaseImponible14D"
Option Explicit

'Forma de ingreso de cada elemento
Public Const ING_MANUAL = 1
Public Const ING_TRASPASO = 2
Public Const ING_TRASPASOAJUSTE = 3
Public Const ING_TRASPASOLIBCAJA = 4
Public Const ING_AMBOS = 5
Public Const ING_AMBOSAJUSTE = 6

'ítems dentro de Ingreso/Egreso
Public Const MAX_BASEIMP14D = 120

'tipo Ingreso o Egreso
Public Const BIMP14D_INGRESO = 1
Public Const BIMP14D_EGRESO = 2

'máx cantidad de niveles
Public Const BIMP14D_MAXNIV = 5


Type DetBaseImponible14D_t
   Tipo As Integer
   Nivel As Integer
   Nombre As String
   Regimen As String
   Codigo As Integer
   FormaIngreso As Integer
   AnoDesde As Integer
   IdItemCtasAsociadasAjustes As Integer
End Type

Public gBaseImponible14D(MAX_BASEIMP14D) As DetBaseImponible14D_t
Public Percepciones(MAX_BASEIMP14D) As DetBaseImponible14D_t

Public Sub InitBaseImponible14D()
   Dim i As Integer
   Dim Codigo As Integer
      
   i = 1
   
   'Base imponible - Codigo HR = 801
      
   gBaseImponible14D(i).Nivel = 1
   gBaseImponible14D(i).Tipo = 0
   gBaseImponible14D(i).Nombre = "Base Imponible"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1
   gBaseImponible14D(i).FormaIngreso = 0
       
      
   'TOTAL INGRESOS
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 2
   gBaseImponible14D(i).Nombre = "Total Ingresos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 100
   gBaseImponible14D(i).FormaIngreso = 0
   
   'Ingresos Percibidos
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Ingresos percibidos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 200
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos percibidos del Giro"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 300
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingresos percibidos del Giro"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 400
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOLIBCAJA
   
   'ADO 2747807 Tema 3 se agrega item
   If gEmpresa.Ano > 2020 Then
    i = i + 1
    gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
    gBaseImponible14D(i).Nivel = 5
    gBaseImponible14D(i).Nombre = "Ingreso del Giro Devengados en ejercicios anteriores y percibidos en el ejercicio actual"
    gBaseImponible14D(i).Regimen = 0   'ambos
    gBaseImponible14D(i).Codigo = 401
    gBaseImponible14D(i).FormaIngreso = ING_TRASPASOLIBCAJA
   End If
   '***************************************
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Renta de Fuente Extranjera"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 500
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos percibidos del art 20 n° 1 LIR"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 600
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Desarrollo de una actividad agrícola"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 700
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 8
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Arriendo de bienes raices agrícolas"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 800
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 9

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Arriendo de bienes raices no agrícolas"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 900
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 10
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos percibidos por la tenencia de capitales mobiliarios (distinto de dividendos)"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1000
   gBaseImponible14D(i).FormaIngreso = 0
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Intereses de depósitos o instrumentos financieros"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1100
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 11
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Mayor valor en el rescate de cuotas de FM o FI"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1200
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 12

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Participación en contratos de participación o cuentas de participación"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1300
   gBaseImponible14D(i).FormaIngreso = ING_AMBOSAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 13
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Retiros o dividendos recibidos desde otras empresas acogidas al regimen 14A, 14B o 14D3"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 1400
   gBaseImponible14D(i).FormaIngreso = 0
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto Liquido recibido"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 1500
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Incremento por crédito IDPC"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 1600
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Incremento por crédito IPE"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 1700
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos percibidos por la enajenación de inversiones"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1800
   gBaseImponible14D(i).FormaIngreso = 0
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Renta percibida en la enajenación de las inversiones en capitales mobiliarios"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 1900
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Renta percibida en la enajenación de contratos de asociación o cuentas en participación"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2000
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Enajenación de Acciones"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Enajenación de derechos sociales"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2200
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Terrenos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2300
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otros  bienes que no se pueden depreciar"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2400
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos percibidos por la enajenación de bienes depreciables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2500
   gBaseImponible14D(i).FormaIngreso = 0
     
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingresos percibidos por la enajenación de bienes depreciables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2600
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 14

   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Reajustes percibidos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2700
   gBaseImponible14D(i).FormaIngreso = 0
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Reajuste de PPM"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2800
   'gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '2699582
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   gBaseImponible14D(i).FormaIngreso = 0
   Else
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   End If
   ' fin 2699582
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Reajuste del remanente de IVA crédito fiscal"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 2900
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Reajustes de depósitos a plazos en UF"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3000
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Reajustes de préstamos en moneda extanjera"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Otros ingresos percibidos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3200
   gBaseImponible14D(i).FormaIngreso = 0
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otros ingresos percibidos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3300
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO


   'Ingresos devengados

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Ingresos devengados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3400
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingresos devengados o percibidos con empresas relacionadas acogidas al régimen 14 Letra A."
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3500
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingresos devengados o percibidos con empresas relacionadas acogidas al régimen 14 Letra A."
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3600
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOLIBCAJA
   
   '------ Nivel 4 ------

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Otros Ingresos Devengados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3700
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otros Ingresos Devengados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3800
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   'Ingresos diferidos

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Ingresos diferidos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 3900
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingreso diferido según art. 14 o 15 transitorio Ley n°21.210"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4000
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingreso diferido incrementado imputado en el año"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '------ Nivel 4 ------

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ingreso diferido para empresas que ingresaron  al régimen de transparencia tributaria"
   gBaseImponible14D(i).Regimen = 0  'FTE_14DN8
   gBaseImponible14D(i).Codigo = 4200
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingreso diferido incrementado imputado en el año"
   gBaseImponible14D(i).Regimen = 0    'FTE_14DN8
   gBaseImponible14D(i).Codigo = 4300
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   gBaseImponible14D(i).AnoDesde = 2021

   'Otros Ajustes que se agregan a la Base Imponible

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Otros Ajustes que se agregan a la Base Imponible"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4400
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Otros Ajustes que se agregan a la Base Imponible"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4500
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Crédito art. 33 Bis LIR"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4600
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Crédito total disponible por IPE"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4700
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_INGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otros agregados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4800
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   
   
   'TOTAL EGRESOS
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 2
   gBaseImponible14D(i).Nombre = "Total Egresos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 4900
   gBaseImponible14D(i).FormaIngreso = 0
   
   'Ingresos Percibidos
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Egresos pagados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5000
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Gastos pagados asociados  en la adquisión de activo realizable o para la prestación de servicios"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5100
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   If gEmpresa.Ano >= 2021 Then
    gBaseImponible14D(i).Nombre = "Existencias, insumos y servicios del negocio, pagados"
   Else
    gBaseImponible14D(i).Nombre = "Existencias o Insumos del Negocio Pagados"
   End If
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5200
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOLIBCAJA
   
   ' INICIO ADO 2747807 Tema 5
   If gEmpresa.Ano >= 2021 Then
    i = i + 1
    gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
    gBaseImponible14D(i).Nivel = 5
    gBaseImponible14D(i).Nombre = "Existencias, insumos y servicios del negocio adeudados en ejercicios anteriores y pagados en el ejercicio actual"
    gBaseImponible14D(i).Regimen = 0   'ambos
    gBaseImponible14D(i).Codigo = 5201
    gBaseImponible14D(i).FormaIngreso = ING_TRASPASOLIBCAJA
   End If
   ' FIN ADO 2747807 Tema 5


    ' ADO 2747807 Tema 1 Servicios pagados solo aparece antes del año 2021
    If gEmpresa.Ano < 2021 Then
       i = i + 1
       gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
       gBaseImponible14D(i).Nivel = 5
       gBaseImponible14D(i).Nombre = "Servicios Pagados"
       gBaseImponible14D(i).Regimen = 0   'ambos
       gBaseImponible14D(i).Codigo = 5300
       gBaseImponible14D(i).FormaIngreso = ING_MANUAL
    End If

   '------ Nivel 4 ------
      

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Montos pagados por la compra de activos fijos depreciables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5400
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Compra de activos fijos depreciables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5500
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO


   '------ Nivel 4 ------
      
         
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Montos pagados asociados a la mantención del giro"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5600
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Remuneraciones pagadas"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5700
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Honorarios pagados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5800
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 4

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Intereses pagados por préstamos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 5900
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Impuestos que no sean de la LIR"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6000
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 5

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Arriendos pagados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6100
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos de Rentas de Fuente Extranjera"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6200
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
    
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Egresos pagados por la adquisición de inversiones o activos en el año de la percepción por la enajenación"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6300
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por inversiones en capitales mobiliarios"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6400
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por contratos de asociación o cuentas en participación"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6500
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por adquisición de acciones"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6600
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por adquisición de derechos sociales"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6700
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por adquisición de terrenos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6800
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Monto pagado reajustado por adquisición de otros bienes que no se deprecian"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 6900
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Gastos afectos al articulo 21 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7000
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos afectos al inciso 1°, del art 21 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7100
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 6

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos afectos al inciso 3°, del art 21 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7200
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 7
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Gastos rechazados pagados no gravados con el art. 21 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7300
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Partidas pagadas del inciso 1° del art 21, no afectas al I.U."
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7400
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago de IDPC"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7500
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 9

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago de IDPC AT 2020 o anteriores que depuran REX"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7600
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 10

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos asociados a INR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7700
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 11

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago 30% ISFUT"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7800
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 12
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otras partidas pagadas del inciso 2° del art 21, distintos de los anteriores"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 7900
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 13


   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Gastos rechazados pagados"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 8000
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos pagados que no cumplen los requisitos del art 31 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 8100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Agregado por gastos rechazados pagados"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 8200
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos pagados que no cumplen los requisitos del art 31 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN8
   gBaseImponible14D(i).Codigo = 8300
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   'Créditos Incobrables por ingresos que ya se reconocieron sobre base devengada
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Créditos Incobrables por ingresos que ya se reconocieron sobre base devengada"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8400
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Créditos incobrables castigados"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8500
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Créditos incobrables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8600
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASOAJUSTE
   gBaseImponible14D(i).IdItemCtasAsociadasAjustes = 8

   'Ajustes en el año por cambio de régimen
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Ajustes en el año por cambio de régimen"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8700
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ajustes en el año por cambio de régimen, al ingresar al régimen del Art. 14D"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8800
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Valor Neto de Activos fijos depreciables"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 8900
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Existencias del activo realizable"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9000
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos diferidos que se mantenian en el activo"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gasto por Pérdida Tributaria en cambio de Régimen"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9200
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   'Otros Gastos que se deducen de la Base Imponible
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Otros Gastos que se deducen de la Base Imponible"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9300
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Otros Gastos que se deducen de la Base Imponible"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9400
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos adeudados asociados a ingresos devengados con empresas relacionadas del régimen 14A"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9500
   'gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   '2699582
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
   
   Else
   
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
    
   End If
   'fin 2699582

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pérdida del año anterior"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9600
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos por responsabilidad social"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 9700
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos por inversión en investigación y desarrollo no certificados por CORFO"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 9800
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos por inversión en investigación y desarrollo certificados por CORFO"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 9900
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   
   '3402617
    If (gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp) And gEmpresa.Ano >= 2023 Then

    Else
        i = i + 1
        gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
        gBaseImponible14D(i).Nivel = 5
        gBaseImponible14D(i).Nombre = "Amortización de intangibles, art. 22° transitorio bis, inc. 4°, 5° y 6° Ley N° 21.210"
        gBaseImponible14D(i).Regimen = 0   'ambos
        gBaseImponible14D(i).Codigo = 10000
        gBaseImponible14D(i).FormaIngreso = ING_MANUAL
    End If
   '3402617
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos aceptados por donaciones"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10100
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingresos Exentos de IDPC"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10200
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Ingresos no rentas"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10300
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otras deducciones a la RLI"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 10400
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
    
     'pipe tema 2 2738156
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otros Gastos"
   gBaseImponible14D(i).Regimen = 0   'ambos
   gBaseImponible14D(i).Codigo = 10410
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
      
   '3402617 ffv
    If (gEmpresa.ProPymeTransp) And gEmpresa.Ano >= 2023 Then
        i = i + 1
        gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
        gBaseImponible14D(i).Nivel = 5
        gBaseImponible14D(i).Nombre = "Gastos por Donaciones"
        gBaseImponible14D(i).Regimen = FTE_14DN8
        gBaseImponible14D(i).Codigo = 10420
        gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
    End If
   '3402617 ffv
   
   'Ajustes a la Base Imponible
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 3
   gBaseImponible14D(i).Nombre = "Ajustes a la Base Imponible"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10500
   gBaseImponible14D(i).FormaIngreso = 0
   
   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ajustes a la Base Imponible"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10600
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Agregado por gastos rechazados pagados"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10700
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Partidas pagadas del inciso 1° del art 21, no afectas al I.U."
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10800
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO
  
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago de IDPC"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 10900
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago de IDPC AT 2020 o anteriores que depuran REX"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11000
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Gastos asociados a INR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11100
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Pago 30% ISFUT"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11200
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Otras partidas pagadas del inciso 2° del art 21, distintos de los anteriores"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11300
   gBaseImponible14D(i).FormaIngreso = ING_TRASPASO

   '------ Nivel 4 ------
      
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 4
   gBaseImponible14D(i).Nombre = "Ajustes art 14 Letra E y art 14 A n°6, ambos de la LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11400
   gBaseImponible14D(i).FormaIngreso = 0
   
   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Franquicia Letra E, Art 14 LIR"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11500
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL

   i = i + 1
   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
   gBaseImponible14D(i).Nivel = 5
   gBaseImponible14D(i).Nombre = "Deducción por pago IDPC Voluntario en años anteriores"
   gBaseImponible14D(i).Regimen = FTE_14DN3
   gBaseImponible14D(i).Codigo = 11600
   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   
   'siguiente codigo debe empezar en 11800
   'codigo 11700 utilizado mas arriba

  
End Sub

Public Sub InitPercepciones()
   Dim i As Integer
   Dim Codigo As Integer
      
   i = 1
  
   Percepciones(i).Nivel = 2
   Percepciones(i).Tipo = 0
   Percepciones(i).Nombre = "PERCEPCIONES"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 100
   Percepciones(i).FormaIngreso = 0
      
   i = i + 1
   Percepciones(i).Nivel = 3
   Percepciones(i).Tipo = 0
   Percepciones(i).Nombre = "Montos afectos a IGC o IA"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 200
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   Percepciones(i).Nivel = 4
   Percepciones(i).Tipo = 0
   Percepciones(i).Nombre = "Suma afectos a IGC o IA"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 300
   Percepciones(i).FormaIngreso = 0
   
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Con crédito a contar del 01.01.2017"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 400
   Percepciones(i).FormaIngreso = 0
   
   'Ingresos Percibidos
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Con crédito hasta el 31.12.2016 (STUT)"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 500
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Con crédito con IDPC Voluntario"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 600
   Percepciones(i).FormaIngreso = 0
   
    i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Sin crédito"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 700
   Percepciones(i).FormaIngreso = 0
   
    i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 3
   Percepciones(i).Nombre = "Rentas exentas e Ingresos no rentas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 800
   Percepciones(i).FormaIngreso = 0
   
   
    i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Rentas con tributación cumplida"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 900
   Percepciones(i).FormaIngreso = 0
   
    i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Rentas provenientes de RAP o EX 14 TER"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1000
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Otras rentas percibidas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1100
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Excesos de distribuciones desproporcionadas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1200
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Utilidades afectas a ISFUT Ley 20.780 y 20.899"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1300
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Utilidades afectas a ISFUT Ley 21.210"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1400
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Rentas Exentas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1500
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Rentas exentas de IGC (Art 11, Ley n° 18.401)"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1600
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Rentas Exentas de IGC e IA"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1700
   Percepciones(i).FormaIngreso = 0
   
   ' NO SE MUESTRA
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "SUMA Ingresos no Renta"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1800
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Ingresos no Renta"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 1900
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 2
   Percepciones(i).Nombre = "Créditos asociados a rentas afectas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2000
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 3
   Percepciones(i).Nombre = "Créditos acumulados desde el 01.01.2017"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2100
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos no sujetos a restitución hasta el 31.12.2019"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2200
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2300
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2400
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos no sujetos a restitución a contar del 01.01.2020"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2500
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2600
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2700
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos sujetos a restitución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2800
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 2900
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3000
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Suma Crédito IPE"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3100
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Crédito IPE"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3200
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos acumulados hasta el 31.12.2016"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3300
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3400
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3500
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Crédito IPE"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3600
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 3
   Percepciones(i).Nombre = "Créditos asociados a rentas exentas"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3700
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos sujetos a restitución a contar del 01.01.2017"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3800
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 3900
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4000
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 4
   Percepciones(i).Nombre = "Créditos hasta el 31.12.2016"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4100
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos sin derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4200
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Créditos con derecho a devolución"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4300
   Percepciones(i).FormaIngreso = 0
   
   'NO MOSTRAR
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 3
   Percepciones(i).Nombre = "Crédito Impuesto adicional Ex art 21 LIR"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4400
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Crédito Impuesto adicional Ex art 21 LIR"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4500
   Percepciones(i).FormaIngreso = 0
   
   i = i + 1
   'Percepciones(i).Tipo = BIMP14D_INGRESO
   Percepciones(i).Nivel = 5
   Percepciones(i).Nombre = "Devolución de Capital Art. 17 n°7 LIR"
   Percepciones(i).Regimen = 0   'ambos
   Percepciones(i).Codigo = 4600
   Percepciones(i).FormaIngreso = 0
   
'   i = i + 1
'   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
'   gBaseImponible14D(i).Nivel = 5
'   gBaseImponible14D(i).Nombre = "Franquicia Letra E, Art 14 LIR"
'   gBaseImponible14D(i).Regimen = FTE_14DN3
'   gBaseImponible14D(i).Codigo = 11500
'   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
'
'   i = i + 1
'   gBaseImponible14D(i).Tipo = BIMP14D_EGRESO
'   gBaseImponible14D(i).Nivel = 5
'   gBaseImponible14D(i).Nombre = "Deducción por pago IDPC Voluntario en años anteriores"
'   gBaseImponible14D(i).Regimen = FTE_14DN3
'   gBaseImponible14D(i).Codigo = 11600
'   gBaseImponible14D(i).FormaIngreso = ING_MANUAL
   
   
   'siguiente codigo debe empezar en 11800
   'codigo 11700 utilizado mas arriba

  
End Sub


Public Function GetPercibidosPagados(ByVal TipoOperCaja As Integer, Optional ByVal Regimen As Integer = 0, Optional ByVal IdLibroCaja As Long = 0) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TmpTbl As String
   Dim Rc As Integer
   Dim IdDoc As Long
   Dim TotPagadoDoc As Double, Afecto As Double, Exento As Double, IVAIrrec As Double
   Dim Dif As Double
   Dim AcumBaseImp As Double, Total As Double
   Dim IniAnoActual, TerAnoActual As Long
   
   GetPercibidosPagados = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)
   
   'creamos una tabla temporal con los Ingresos/Egresos y el total pagado/percibido
   Q1 = "SELECT LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, Sum(LibroCaja.Pagado) As TotPagado, LibroCaja.FechaOperacion "
   '2690461
   Q1 = Q1 & ",(select sum(afecto) as monto from documento where tipoDoc = 3 "

   Q1 = Q1 & "and TipoLib = LibroCaja.TipoLib and iddocasoc = LibroCaja.IdDoc) as AfectoNotCred "
 
   'fin 2690461
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "Documento", True, True) & ")"
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad "
   Else
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   End If
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa "
'   Q1 = Q1 & " WHERE Year(Documento.FEmisionOri) = " & gEmpresa.Ano '' FPG 29-10-2021 SOLICITADO POR VICTOR (TK 2670511)
   Q1 = Q1 & " WHERE TipoOper = " & TipoOperCaja & " AND Pagado <> 0 "
   
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.Giro <> 0 "
      
      If Regimen = FTE_14DN3 Or Regimen = FTE_14DN8 Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0 OR ( Entidades.EntRelacionada <> 0 AND FranqTribEnt IN (" & FTE_14DN3 & "," & FTE_14DN8 & ") ) )"
      ElseIf Regimen = FTE_14A Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada <> 0 AND FranqTribEnt = " & FTE_14A & " )"
      End If
   
   Else
      Q1 = Q1 & " AND LibroCaja.TipoLib IN ( " & LIB_COMPRAS & ", " & LIB_RETEN & " ) "
   End If
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   If IdLibroCaja <> 0 Then
      Q1 = Q1 & " AND LibroCaja.IdLibroCaja = " & IdLibroCaja
   End If
   
   Q1 = Q1 & " GROUP BY LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, LibroCaja.FechaOperacion "
   
   Rc = ExecSQL(DbMain, Q1)
   
   TotPagadoDoc = 0
   Dif = 0
   
'   'si es sólo un documento, obtenemos el total pagado/percibido   ESTO NO SE HACE PORQUE RESTA LA CANTIDAD EN TODOS LOS REGISTROS DEL LIBRO DE CAJA ASOCIADOS A ESTE DOCUMENTO
   If IdLibroCaja > 0 Then

      Q1 = "SELECT IdDoc, Afecto, Exento, IVAIrrec FROM LibroCaja WHERE IdLibroCaja = " & IdLibroCaja
      Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdDoc = vFld(Rs("IdDoc"))
         Afecto = vFld(Rs("Afecto"))
         Exento = vFld(Rs("Exento"))
         IVAIrrec = vFld(Rs("IVAIrrec"))
      End If
      Call CloseRs(Rs)

      If IdDoc > 0 Then
         Q1 = "SELECT Sum(LibroCaja.Pagado) As TotPagadoDoc, Sum(MontoAfectaBaseImp) As AcumBaseImp FROM LibroCaja WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdLibroCaja <> " & IdLibroCaja    'sin considerar el registro de libro de caja actual
         Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            TotPagadoDoc = vFld(Rs("TotPagadoDoc"))    'lo pagado sin considerar este registro
            AcumBaseImp = vFld(Rs("AcumBaseImp"))      'base imponoble acumulada sin considerar este registro
         End If
         Call CloseRs(Rs)
         Dif = Afecto + Exento + IVAIrrec - AcumBaseImp
         If Dif < 0 Then
            Dif = 0
         End If
      End If
   End If

   
   'y ahora obtenemos el total de los
   'Ingresos percibidos del Giro (cuando TipoOperCaja = INGRESO)y
   'Existencias o Insumos del Negocio Pagados  ((cuando TipoOperCaja = EGRESO)
   IniAnoActual = CLng(DateSerial(gEmpresa.Ano, 1, 1))
   TerAnoActual = CLng(DateSerial(gEmpresa.Ano, 12, 31))
   
   If IdLibroCaja > 0 Then
      Q1 = "SELECT sum( iif( Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec) < TotPagado, Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec), TotPagado )) As Total"
      
      
      Q1 = Q1 & " FROM " & TmpTbl
      
'      Q1 = "SELECT sum( iif( TMP.Afecto + TMP.Exento + iif(TMP.IVAIrrec IS NULL, 0, TMP.IVAIrrec) < TMP.TotPagado, TMP.Afecto + TMP.Exento + iif(TMP.IVAIrrec IS NULL, 0, TMP.IVAIrrec), TMP.TotPagado )) As Total"
'      Q1 = Q1 & " FROM " & TmpTbl & " AS TMP, DOCUMENTO AS DOC, CUENTAS AS CU "
'      Q1 = Q1 & " WHERE TMP.NumDoc = DOC.NumDoc "
'      Q1 = Q1 & " AND CU.IDCUENTA = DOC.IdCuentaTotal "
'      Q1 = Q1 & " AND IIF(CODF22_14TER IS NULL,0,CODF22_14TER) NOT IN (1413,1618) "
      ' ADO 2747807 Tema 2 se agregar el filtro de fecha para que sea solo del año actual
      Q1 = Q1 & " WHERE Fechaoperacion BETWEEN " & IniAnoActual & " AND " & TerAnoActual & ""
   Else
      'Q1 = "SELECT sum( iif( Afecto + Exento < TotPagado, Afecto + Exento, TotPagado )) As Total"
      
      '2690461
      Q1 = "SELECT sum(iif(AfectoNotCred<> null, iif(Afecto + Exento - AfectoNotCred < TotPagado , Afecto + Exento - AfectoNotCred,TotPagado) , iif( Afecto + Exento < TotPagado, Afecto + Exento, TotPagado )  )) As Total "
      'fin 2690461
      
      Q1 = Q1 & " FROM " & TmpTbl
      Q1 = Q1 & " WHERE TIPOLIB <> 3 "
      
'      Q1 = "SELECT sum( iif( TMP.Afecto + TMP.Exento < TMP.TotPagado, TMP.Afecto + TMP.Exento, TMP.TotPagado )) As Total"
'      Q1 = Q1 & " FROM " & TmpTbl & " AS TMP, DOCUMENTO AS DOC, CUENTAS AS CU "
'      Q1 = Q1 & " WHERE TMP.NumDoc = DOC.NumDoc "
'      Q1 = Q1 & " AND CU.IDCUENTA = DOC.IdCuentaTotal "
'      Q1 = Q1 & " AND IIF(CODF22_14TER IS NULL,0,CODF22_14TER) NOT IN (1413,1618) "
'      Q1 = Q1 & " AND TMP.TIPOLIB <> 3 "
      ' ADO 2747807 Tema 2 se agregar el filtro de fecha para que sea solo del año actual
      Q1 = Q1 & " AND Fechaoperacion BETWEEN " & IniAnoActual & " AND " & TerAnoActual & ""
   End If
   
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = vFld(Rs("Total"))
      
      If IdLibroCaja > 0 Then
         GetPercibidosPagados = IIf(Total > Dif, Dif, Total)
      Else
         GetPercibidosPagados = Total
      End If
      
   End If
   
   Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function


Public Function GetPercibidosPagadosCompras(ByVal TipoOperCaja As Integer, Optional ByVal Regimen As Integer = 0, Optional ByVal IdLibroCaja As Long = 0) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TmpTbl As String
   Dim Rc As Integer
   Dim IdDoc As Long
   Dim TotPagadoDoc As Double, Afecto As Double, Exento As Double, IVAIrrec As Double
   Dim Dif As Double
   Dim AcumBaseImp As Double, Total As Double
   Dim IniAnoActual, TerAnoActual As Long
   
   GetPercibidosPagadosCompras = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)
   
   'creamos una tabla temporal con los Ingresos/Egresos y el total pagado/percibido
   Q1 = "SELECT LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri,LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, Sum(LibroCaja.pagado) As TotPagado, LibroCaja.FechaOperacion   "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "Documento", True, True) & ")"
   If TipoOperCaja = TOPERCAJA_EGRESO Then
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad "
   Else
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   End If
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa "
'   Q1 = Q1 & " WHERE Year(Documento.FEmisionOri) = " & gEmpresa.Ano '' FPG 29-10-2021 SOLICITADO POR VICTOR (TK 2670511)
   Q1 = Q1 & " WHERE TipoOper = " & TipoOperCaja
   Q1 = Q1 & " AND LibroCaja.TipoLib = " & LIB_COMPRAS
   Q1 = Q1 & " AND LibroCaja.estado in  (" & ED_APROBADO & "," & ED_PENDIENTE & "," & ED_CENTRALIZADO & ")"
   Q1 = Q1 & " AND LibroCaja.ConEntRel = -1 "
   If TipoOperCaja = TOPERCAJA_EGRESO Then
      Q1 = Q1 & " AND Documento.Giro <> 0 "
      
     If Regimen = FTE_14A Then
         Q1 = Q1 & " AND FranqTribEnt = " & FTE_14A
      End If

   End If
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   If IdLibroCaja <> 0 Then
      Q1 = Q1 & " AND LibroCaja.IdLibroCaja = " & IdLibroCaja
   End If
   
   Q1 = Q1 & " GROUP BY LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, LibroCaja.FechaOperacion "
   
   Rc = ExecSQL(DbMain, Q1)
   
   TotPagadoDoc = 0
   Dif = 0
      
   'y ahora obtenemos el total de los
   'Ingresos percibidos del Giro (cuando TipoOperCaja = INGRESO)y
   'Existencias o Insumos del Negocio Pagados  ((cuando TipoOperCaja = EGRESO)
   IniAnoActual = CLng(DateSerial(gEmpresa.Ano, 1, 1))
   TerAnoActual = CLng(DateSerial(gEmpresa.Ano, 12, 31))
     
      Q1 = "SELECT sum(afecto + exento)  As Total"
      Q1 = Q1 & " FROM " & TmpTbl
      Q1 = Q1 & " WHERE Fechaoperacion BETWEEN " & IniAnoActual & " AND " & TerAnoActual & ""
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = vFld(Rs("Total"))
      
      If IdLibroCaja > 0 Then
         GetPercibidosPagadosCompras = IIf(Total > Dif, Dif, Total)
      Else
         GetPercibidosPagadosCompras = Total
      End If
      
   End If
   
   Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function


Public Function ExistOInsumPagados(ByVal TipoOperCaja As Integer, Optional ByVal AnoActual As Boolean = True, Optional ByVal Regimen As Integer = 0, Optional ByVal IdLibroCaja As Long = 0) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TmpTbl As String
   Dim Rc As Integer
   Dim IdDoc As Long
   Dim TotPagadoDoc As Double, Afecto As Double, Exento As Double, IVAIrrec As Double
   Dim Dif As Double
   Dim AcumBaseImp As Double, Total As Double
   Dim IniAnoActual, TerAnoActual As Long
   
   ExistOInsumPagados = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)
   
   'creamos una tabla temporal con los Ingresos/Egresos y el total pagado/percibido
   Q1 = "SELECT LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, Sum(LibroCaja.Pagado) As TotPagado, LibroCaja.FechaOperacion "
    '2690461
   Q1 = Q1 & ",(select sum(afecto) as monto from documento where tipoDoc = 3 "

   Q1 = Q1 & "and TipoLib = LibroCaja.TipoLib and iddocasoc = LibroCaja.IdDoc) as AfectoNotCred "
 
   'fin 2690461
   
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "Documento", True, True) & ")"
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad "
   Else
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   End If
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa "
'   Q1 = Q1 & " WHERE Year(Documento.FEmisionOri) = " & gEmpresa.Ano '' FPG 29-10-2021 SOLICITADO POR VICTOR (TK 2670511)
   Q1 = Q1 & " WHERE TipoOper = " & TipoOperCaja & " AND Pagado <> 0 "
   
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.Giro <> 0 "
      
      If Regimen = FTE_14DN3 Or Regimen = FTE_14DN8 Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0 OR ( Entidades.EntRelacionada <> 0 AND FranqTribEnt IN (" & FTE_14DN3 & "," & FTE_14DN8 & ") ) )"
      ElseIf Regimen = FTE_14A Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada <> 0 AND FranqTribEnt = " & FTE_14A & " )"
      End If
   
   Else
      Q1 = Q1 & " AND LibroCaja.TipoLib IN ( " & LIB_COMPRAS & ", " & LIB_RETEN & " ) "
   End If
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   If IdLibroCaja <> 0 Then
      Q1 = Q1 & " AND LibroCaja.IdLibroCaja = " & IdLibroCaja
   End If
   
   Q1 = Q1 & " GROUP BY LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, LibroCaja.FechaOperacion "
   
   Rc = ExecSQL(DbMain, Q1)
   
   TotPagadoDoc = 0
   Dif = 0
   
'   'si es sólo un documento, obtenemos el total pagado/percibido   ESTO NO SE HACE PORQUE RESTA LA CANTIDAD EN TODOS LOS REGISTROS DEL LIBRO DE CAJA ASOCIADOS A ESTE DOCUMENTO
   If IdLibroCaja > 0 Then

      Q1 = "SELECT IdDoc, Afecto, Exento, IVAIrrec FROM LibroCaja WHERE IdLibroCaja = " & IdLibroCaja
      Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdDoc = vFld(Rs("IdDoc"))
         Afecto = vFld(Rs("Afecto"))
         Exento = vFld(Rs("Exento"))
         IVAIrrec = vFld(Rs("IVAIrrec"))
      End If
      Call CloseRs(Rs)

      If IdDoc > 0 Then
         Q1 = "SELECT Sum(LibroCaja.Pagado) As TotPagadoDoc, Sum(MontoAfectaBaseImp) As AcumBaseImp FROM LibroCaja WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdLibroCaja <> " & IdLibroCaja    'sin considerar el registro de libro de caja actual
         Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            TotPagadoDoc = vFld(Rs("TotPagadoDoc"))    'lo pagado sin considerar este registro
            AcumBaseImp = vFld(Rs("AcumBaseImp"))      'base imponoble acumulada sin considerar este registro
         End If
         Call CloseRs(Rs)
         Dif = Afecto + Exento + IVAIrrec - AcumBaseImp
         If Dif < 0 Then
            Dif = 0
         End If
      End If
   End If

   
   'y ahora obtenemos el total de los
   'Ingresos percibidos del Giro (cuando TipoOperCaja = INGRESO)y
   'Existencias o Insumos del Negocio Pagados  ((cuando TipoOperCaja = EGRESO)
   IniAnoActual = CLng(DateSerial(gEmpresa.Ano, 1, 1))
   TerAnoActual = CLng(DateSerial(gEmpresa.Ano, 12, 31))
   
   If IdLibroCaja > 0 Then
'      Q1 = "SELECT sum( iif( Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec) < TotPagado, Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec), TotPagado )) As Total"
'      Q1 = Q1 & " FROM " & TmpTbl
      
      Q1 = "SELECT sum( iif( " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento + iif(" & TmpTbl & ".IVAIrrec IS NULL, 0, " & TmpTbl & ".IVAIrrec) < " & TmpTbl & ".TotPagado, " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento + iif(" & TmpTbl & ".IVAIrrec IS NULL, 0, " & TmpTbl & ".IVAIrrec), " & TmpTbl & ".TotPagado )) As Total"
      Q1 = Q1 & " FROM (" & TmpTbl & " INNER JOIN DOCUMENTO      "
       'Q1 = Q1 & " ON  " & TmpTbl & ".numdoc = DOCUMENTO.numdoc)"
       Q1 = Q1 & " ON  " & TmpTbl & ".iddoc = DOCUMENTO.iddoc)" '2782352
      Q1 = Q1 & " LEFT JOIN CUENTAS ON CUENTAS.IDCUENTA = DOCUMENTO.IdCuentaTotal"
      Q1 = Q1 & " WHERE   IIF(CODF22_14TER IS NULL,0,CODF22_14TER) NOT IN (1413,1618)"
      ' ADO 2747807 Tema 4 se agregar el filtro de fecha para que sea solo del año actual
      ' ADO 2747807 Tema 5 (else) se agregar el filtro de fecha para que sea solo de años anteriores
      If gEmpresa.Ano >= 2021 Then
          If AnoActual Then
            Q1 = Q1 & " AND Fechaoperacion BETWEEN " & IniAnoActual & " AND " & TerAnoActual & ""
          Else
            Q1 = Q1 & " AND Fechaoperacion < " & IniAnoActual & ""
          End If
      End If
   Else
'      Q1 = "SELECT sum( iif( Afecto + Exento < TotPagado, Afecto + Exento, TotPagado )) As Total"
'      Q1 = Q1 & " FROM " & TmpTbl
'      Q1 = Q1 & " WHERE TIPOLIB <> 3 "
      
      'Q1 = "SELECT sum( iif( " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento < " & TmpTbl & ".TotPagado, " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento, " & TmpTbl & ".TotPagado )) As Total"
      
    '2690461
      Q1 = "SELECT sum(iif(AfectoNotCred<> null, iif( " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento - AfectoNotCred < " & TmpTbl & ".TotPagado, " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento - AfectoNotCred , " & TmpTbl & ".TotPagado ), iif(" & TmpTbl & ".Afecto + " & TmpTbl & ".Exento < " & TmpTbl & ".TotPagado," & TmpTbl & ".Afecto + " & TmpTbl & ".Exento," & TmpTbl & ".TotPagado )  )) As Total "
      'fin 2690461
      
      'Q1 = "SELECT sum(iif(AfectoNotCred<> null, iif( " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento - AfectoNotCred < " & TmpTbl & ".TotPagado, " & TmpTbl & ".Afecto + " & TmpTbl & ".Exento - AfectoNotCred , " & TmpTbl & ".TotPagado ), iif( Afecto + Exento < TotPagado, Afecto + Exento, TotPagado )  )) As Total"
      
      Q1 = Q1 & " FROM (" & TmpTbl & " INNER JOIN DOCUMENTO "
      'Q1 = Q1 & " ON  " & TmpTbl & ".numdoc = DOCUMENTO.numdoc)"
       Q1 = Q1 & " ON  " & TmpTbl & ".iddoc = DOCUMENTO.iddoc)" '2782352
      Q1 = Q1 & " LEFT JOIN CUENTAS ON CUENTAS.IDCUENTA = DOCUMENTO.IdCuentaTotal"
      Q1 = Q1 & " WHERE   IIF(CODF22_14TER IS NULL,0,CODF22_14TER) NOT IN (1413,1618)"
      Q1 = Q1 & " AND " & TmpTbl & ".TIPOLIB <> 3 "
'       ADO 2747807 Tema 4 se agregar el filtro de fecha para que sea solo del año actual
'       ADO 2747807 Tema 5 (else) se agregar el filtro de fecha para que sea solo de años anteriores
      If gEmpresa.Ano >= 2021 Then
          If AnoActual Then
            Q1 = Q1 & " AND Fechaoperacion BETWEEN " & IniAnoActual & " AND " & TerAnoActual & ""
          Else
            Q1 = Q1 & " AND Fechaoperacion < " & IniAnoActual & ""
          End If
      End If
   End If
   
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = vFld(Rs("Total"))
      
      If IdLibroCaja > 0 Then
         ExistOInsumPagados = IIf(Total > Dif, Dif, Total)
      Else
         ExistOInsumPagados = Total
      End If
      
   End If
   
   Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function

Public Function GetActivosFijosDepreciables(ByVal TipoOperCaja As Integer, Optional ByVal Regimen As Integer = 0, Optional ByVal IdLibroCaja As Long = 0) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TmpTbl As String
   Dim Rc As Integer
   Dim IdDoc As Long
   Dim TotPagadoDoc As Double, Afecto As Double, Exento As Double, IVAIrrec As Double
   Dim Dif As Double
   Dim AcumBaseImp As Double, Total As Double
   Dim IniAnoActual, TerAnoActual As Long
   
   '2804908
   Dim Q2 As String
   Dim Rs2 As Recordset
   'FIN 2804908
   
   GetActivosFijosDepreciables = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)
   
   'creamos una tabla temporal con los Ingresos/Egresos y el total pagado/percibido
   Q1 = "SELECT LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, Sum(LibroCaja.Pagado) As TotPagado, LibroCaja.FechaOperacion "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "Documento", True, True) & ")"
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad "
   Else
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   End If
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa "
'   Q1 = Q1 & " WHERE Year(Documento.FEmisionOri) = " & gEmpresa.Ano '' FPG 29-10-2021 SOLICITADO POR VICTOR (TK 2670511)
   
   Q1 = Q1 & " WHERE TipoOper = " & TipoOperCaja & " AND Pagado <> 0 "
   
   'Q1 = Q1 & " WHERE TipoOper = 1 AND Pagado <> 0 "
     
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.Giro <> 0 "
      
      If Regimen = FTE_14DN3 Or Regimen = FTE_14DN8 Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0 OR ( Entidades.EntRelacionada <> 0 AND FranqTribEnt IN (" & FTE_14DN3 & "," & FTE_14DN8 & ") ) )"
      ElseIf Regimen = FTE_14A Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada <> 0 AND FranqTribEnt = " & FTE_14A & " )"
      End If
   
   Else
      Q1 = Q1 & " AND LibroCaja.TipoLib IN ( " & LIB_COMPRAS & ", " & LIB_RETEN & " ) "
    'Q1 = Q1 & " AND LibroCaja.TipoLib IN ( 2 ) "
   End If
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   If IdLibroCaja <> 0 Then
      Q1 = Q1 & " AND LibroCaja.IdLibroCaja = " & IdLibroCaja
   End If
   
   Q1 = Q1 & " GROUP BY LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, LibroCaja.FechaOperacion "
   
   Rc = ExecSQL(DbMain, Q1)
   
   TotPagadoDoc = 0
   Dif = 0
   
'   'si es sólo un documento, obtenemos el total pagado/percibido   ESTO NO SE HACE PORQUE RESTA LA CANTIDAD EN TODOS LOS REGISTROS DEL LIBRO DE CAJA ASOCIADOS A ESTE DOCUMENTO
   If IdLibroCaja > 0 Then

      Q1 = "SELECT IdDoc, Afecto, Exento, IVAIrrec FROM LibroCaja WHERE IdLibroCaja = " & IdLibroCaja
      Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdDoc = vFld(Rs("IdDoc"))
         Afecto = vFld(Rs("Afecto"))
         Exento = vFld(Rs("Exento"))
         IVAIrrec = vFld(Rs("IVAIrrec"))
      End If
      Call CloseRs(Rs)

      If IdDoc > 0 Then
         Q1 = "SELECT Sum(LibroCaja.Pagado) As TotPagadoDoc, Sum(MontoAfectaBaseImp) As AcumBaseImp FROM LibroCaja WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdLibroCaja <> " & IdLibroCaja    'sin considerar el registro de libro de caja actual
         Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            TotPagadoDoc = vFld(Rs("TotPagadoDoc"))    'lo pagado sin considerar este registro
            AcumBaseImp = vFld(Rs("AcumBaseImp"))      'base imponoble acumulada sin considerar este registro
         End If
         Call CloseRs(Rs)
         Dif = Afecto + Exento + IVAIrrec - AcumBaseImp
         If Dif < 0 Then
            Dif = 0
         End If
      End If
   End If

   
   'y ahora obtenemos el total de los
   'Ingresos percibidos del Giro (cuando TipoOperCaja = INGRESO)y
   'Existencias o Insumos del Negocio Pagados  ((cuando TipoOperCaja = EGRESO)
   IniAnoActual = CLng(DateSerial(gEmpresa.Ano, 1, 1))
   TerAnoActual = CLng(DateSerial(gEmpresa.Ano, 12, 31))
   
   If IdLibroCaja > 0 Then
      Q1 = "SELECT sum( iif( TMP.Afecto + TMP.Exento + iif(TMP.IVAIrrec IS NULL, 0, TMP.IVAIrrec) < TMP.TotPagado, TMP.Afecto + TMP.Exento + iif(TMP.IVAIrrec IS NULL, 0, TMP.IVAIrrec), TMP.TotPagado )) As Total"
      Q1 = Q1 & " FROM " & TmpTbl & " AS TMP, DOCUMENTO AS DOC, CUENTAS AS CU "
      Q1 = Q1 & " WHERE TMP.NumDoc = DOC.NumDoc "
      Q1 = Q1 & " AND CU.IDCUENTA = DOC.IdCuentaTotal "
      Q1 = Q1 & " AND CODF22_14TER in (1413,1618) "

   Else
      Q1 = "SELECT sum( iif( TMP.Afecto + TMP.Exento < TMP.TotPagado, TMP.Afecto + TMP.Exento, TMP.TotPagado )) As Total"
      Q1 = Q1 & " FROM " & TmpTbl & " AS TMP, DOCUMENTO AS DOC, CUENTAS AS CU "
      Q1 = Q1 & " WHERE TMP.NumDoc = DOC.NumDoc "
      Q1 = Q1 & " AND CU.IDCUENTA = DOC.IdCuentaTotal "
      Q1 = Q1 & " AND CODF22_14TER in (1413,1618) "
      Q1 = Q1 & " AND TMP.TIPOLIB <> 3 "
   End If
   
   '
   '2804908
      Q2 = ""
      Q2 = " SELECT sum(iif(MovComprobante.debe>0, MovComprobante.debe,MovComprobante.haber)) As Total"
      Q2 = Q2 & " FROM Cuentas INNER JOIN (Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp) ON Cuentas.idCuenta = MovComprobante.IdCuenta "
      Q2 = Q2 & " WHERE CODF22_14TER in (1413,1618) AND COMPROBANTE.ESTADO = 2 AND TIPO = 1 AND IDDOC = 0 "
        
        Set Rs2 = OpenRs(DbMain, Q2)
   
   If Not Rs2.EOF Then
      Total = vFld(Rs2("Total"))
   End If
   
   Call CloseRs(Rs2)
   
   'FIN 2804908
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = Total + vFld(Rs("Total"))
      
      If IdLibroCaja > 0 Then
         GetActivosFijosDepreciables = IIf(Total > Dif, Dif, Total)
      Else
         GetActivosFijosDepreciables = Total
      End If
      
   End If
   
   Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function
' 2810388
Public Function GetNCVParaExistencias() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Total As Double
   GetNCVParaExistencias = 0
   
   
   'If IdLibroCaja > 0 Then
      Q1 = "SELECT sum( iif( LI.Afecto + LI.Exento < LI.Pagado, LI.Afecto + LI.Exento, LI.Pagado )) As Total "
      Q1 = Q1 & " FROM COMPROBANTE CO, MOVCOMPROBANTE MO, DOCUMENTO DO, LIBROCAJA LI "
      Q1 = Q1 & " Where CO.IdComp = MO.IdComp "
      Q1 = Q1 & " AND MO.IDDOC = DO.IDDOC "
      Q1 = Q1 & " AND DO.IDDOC = LI.IDDOC "
      Q1 = Q1 & " AND CO.TIPO = 1 "
      '2839834
      Q1 = Q1 & " AND CO.ESTADO in (2,3) "
      'Q1 = Q1 & " AND CO.ESTADO = 2 "
      'fin 2839834
      Q1 = Q1 & " AND DO.TIPOLIB = 2 "
      Q1 = Q1 & " AND DO.TIPODOC = 3 "
      Q1 = Q1 & " AND DO.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " AND DO.ano = " & gEmpresa.Ano
      
   'End If
   
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = vFld(Rs("Total"))
      GetNCVParaExistencias = Total
   End If
   
   Call CloseRs(Rs)


End Function


'ADO 2747807 Tema 3 Se agrega metodo GetDevengadosAnteriores para el calculo de años anteriores
Public Function GetDevengadosAnteriores(ByVal TipoOperCaja As Integer, Optional ByVal Regimen As Integer = 0, Optional ByVal IdLibroCaja As Long = 0) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TmpTbl As String
   Dim Rc As Integer
   Dim IdDoc As Long
   Dim TotPagadoDoc As Double, Afecto As Double, Exento As Double, IVAIrrec As Double
   Dim Dif As Double
   Dim AcumBaseImp As Double, Total As Double
   Dim IniAnoActual, TerAnoActual As Long
   
   GetDevengadosAnteriores = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)
   
   'creamos una tabla temporal con los Ingresos/Egresos y el total pagado/percibido
   Q1 = "SELECT LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, Sum(LibroCaja.Pagado) As TotPagado, LibroCaja.FechaOperacion "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM (LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "LibroCaja", "Documento", True, True) & ")"
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad "
   Else
      Q1 = Q1 & " LEFT JOIN Entidades ON LibroCaja.IdEntidad = Entidades.IdEntidad "
   End If
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = Entidades.IdEmpresa "
'   Q1 = Q1 & " WHERE Year(Documento.FEmisionOri) = " & gEmpresa.Ano '' FPG 29-10-2021 SOLICITADO POR VICTOR (TK 2670511)
   Q1 = Q1 & " WHERE TipoOper = " & TipoOperCaja & " AND Pagado <> 0 "
   
   If TipoOperCaja = TOPERCAJA_INGRESO Then
      Q1 = Q1 & " AND Documento.Giro <> 0 "
      
      If Regimen = FTE_14DN3 Or Regimen = FTE_14DN8 Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada IS NULL OR Entidades.EntRelacionada = 0 OR ( Entidades.EntRelacionada <> 0 AND FranqTribEnt IN (" & FTE_14DN3 & "," & FTE_14DN8 & ") ) )"
      ElseIf Regimen = FTE_14A Then
         Q1 = Q1 & " AND ( Entidades.EntRelacionada <> 0 AND FranqTribEnt = " & FTE_14A & " )"
      End If
   
   Else
      Q1 = Q1 & " AND LibroCaja.TipoLib IN ( " & LIB_COMPRAS & ", " & LIB_RETEN & " ) "
   End If
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   
   If IdLibroCaja <> 0 Then
      Q1 = Q1 & " AND LibroCaja.IdLibroCaja = " & IdLibroCaja
   End If
   
   Q1 = Q1 & " GROUP BY LibroCaja.IdDoc, LibroCaja.TipoOper, LibroCaja.TipoLib, LibroCaja.TipoDoc, LibroCaja.DTE, LibroCaja.NumDoc, Documento.FEmisionOri, LibroCaja.Afecto, LibroCaja.Exento, LibroCaja.IVAIrrec, LibroCaja.FechaOperacion "
   
   Rc = ExecSQL(DbMain, Q1)
   
   TotPagadoDoc = 0
   Dif = 0
   
'   'si es sólo un documento, obtenemos el total pagado/percibido   ESTO NO SE HACE PORQUE RESTA LA CANTIDAD EN TODOS LOS REGISTROS DEL LIBRO DE CAJA ASOCIADOS A ESTE DOCUMENTO
   If IdLibroCaja > 0 Then

      Q1 = "SELECT IdDoc, Afecto, Exento, IVAIrrec FROM LibroCaja WHERE IdLibroCaja = " & IdLibroCaja
      Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdDoc = vFld(Rs("IdDoc"))
         Afecto = vFld(Rs("Afecto"))
         Exento = vFld(Rs("Exento"))
         IVAIrrec = vFld(Rs("IVAIrrec"))
      End If
      Call CloseRs(Rs)

      If IdDoc > 0 Then
         Q1 = "SELECT Sum(LibroCaja.Pagado) As TotPagadoDoc, Sum(MontoAfectaBaseImp) As AcumBaseImp FROM LibroCaja WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdLibroCaja <> " & IdLibroCaja    'sin considerar el registro de libro de caja actual
         Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            TotPagadoDoc = vFld(Rs("TotPagadoDoc"))    'lo pagado sin considerar este registro
            AcumBaseImp = vFld(Rs("AcumBaseImp"))      'base imponoble acumulada sin considerar este registro
         End If
         Call CloseRs(Rs)
         Dif = Afecto + Exento + IVAIrrec - AcumBaseImp
         If Dif < 0 Then
            Dif = 0
         End If
      End If
   End If

   
   'y ahora obtenemos el total de los
   'Ingreso del Giro Devengados en ejercicios anteriores y percibidos en el ejercicio actual (cuando TipoOperCaja = INGRESO)y
   'Existencias o Insumos del Negocio Pagados  ((cuando TipoOperCaja = EGRESO)
   IniAnoActual = CLng(DateSerial(gEmpresa.Ano, 1, 1))
   
   If IdLibroCaja > 0 Then
      Q1 = "SELECT sum( iif( Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec) < TotPagado, Afecto + Exento + iif(IVAIrrec IS NULL, 0, IVAIrrec), TotPagado )) As Total"
      Q1 = Q1 & " FROM " & TmpTbl
      Q1 = Q1 & " WHERE Fechaoperacion < " & IniAnoActual & ""
   Else
      Q1 = "SELECT sum( iif( Afecto + Exento < TotPagado, Afecto + Exento, TotPagado )) As Total"
      Q1 = Q1 & " FROM " & TmpTbl
      Q1 = Q1 & " WHERE TIPOLIB <> 3 "
      Q1 = Q1 & " AND Fechaoperacion < " & IniAnoActual & ""
   End If
   
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Total = vFld(Rs("Total"))
      
      If IdLibroCaja > 0 Then
         GetDevengadosAnteriores = IIf(Total > Dif, Dif, Total)
      Else
         GetDevengadosAnteriores = Total
      End If
      
   End If
   
   Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function




'ADO 2699582 Tema 3.2 Se agrega metodo GetPerdidaAnoAnterior para el calculo de saldos años anteriores
Public Function GetPerdidaAnoAnterior() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
    Dim Q3 As String
   Dim Rs3 As Recordset
   

   Dim Total As Double
   Dim IniAnoAnterior As Long

   GetPerdidaAnoAnterior = 0

   IniAnoAnterior = CLng(gEmpresa.Ano - 1)
      
            Q1 = "SELECT sum(valor) as Total FROM BaseImponible14D"
              Q1 = Q1 & " WHERE ano = " & IniAnoAnterior
            Q1 = Q1 & " AND FECHA = 0 "
        
        
           Set Rs = OpenRs(DbMain, Q1)
        
           If Not Rs.EOF Then
              Total = vFld(Rs("Total"))
              GetPerdidaAnoAnterior = Total
              
                           
           Q2 = "SELECT sum(valor) as Total FROM BaseImponible14D"
            Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
            Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0"
                
        
           Set Rs2 = OpenRs(DbMain, Q2)
        
           If Not Rs2.EOF Then
              Total = vFld(Rs2("Total"))
        
              If Total > 0 Then
               GetPerdidaAnoAnterior = Total * -1
                
              Else
                    Q3 = "SELECT valor as Total FROM BaseImponible14D"
                    Q3 = Q3 & " WHERE ano = " & gEmpresa.Ano
                    Q3 = Q3 & " AND CODIGO = 9600 AND FECHA = 0 and valor = 0"
                
                
                   Set Rs3 = OpenRs(DbMain, Q3)
                
                   If Not Rs3.EOF Then
                       Total = vFld(Rs3("Total"))
                
                      If Total = 0 Then
                      GetPerdidaAnoAnterior = Total
                        Call CloseRs(Rs)
                        Call CloseRs(Rs2)
                        Call CloseRs(Rs3)
                      Exit Function
                      
                      End If
                  End If
              End If
           End If
              
           End If
   
  

   Call CloseRs(Rs)
   Call CloseRs(Rs2)
   Call CloseRs(Rs3)


End Function

'2699582
Public Sub TraerPerdidaAnterior(lTipo As Integer)
    Dim Q1 As String
    Dim Rs As Recordset
    Dim Q2 As String
    Dim Rs2 As Recordset
    Dim Q3 As String
    Dim Rs3 As Recordset
   
    Dim DbName As String
    Dim DbAnoAnt As Database
    Dim Total As Double
    Dim Fecha As Long
    Dim IniAnoAnterior As Long
    
    Dim Rs5 As dao.Recordset
    Dim ConnStr As String
    Dim trae As Boolean

 IniAnoAnterior = CLng(gEmpresa.Ano - 1)
 
 Fecha = GetDate("3112" & gEmpresa.Ano, "dmy")
 
#If DATACON = 1 Then
'If gDbType = SQL_ACCESS Then
 
    If gEmpresa.TieneAnoAnt Then
       DbName = Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\..\", "")
       If ExistFile(DbName) Then
          Call OpenDbEmp2(DbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
       Else
          MsgBox1 "No Existe base de datos del año anterior", vbInformation
       End If
    End If
    
     Q1 = "SELECT sum(valor) as Total FROM BaseImponible14D"
            Q1 = Q1 & " WHERE ano = " & IniAnoAnterior
            Q1 = Q1 & " AND FECHA = 0 "
        
          Set Rs = OpenRs(DbAnoAnt, Q1)

      If Not Rs.EOF Then
              Total = vFld(Rs("Total"))
              
              If Total >= 0 Then
              Call CloseRs(Rs)
               Call CloseDb(DbAnoAnt)
                MsgBox1 "No existe Perdida del Año Anterior", vbExclamation
              Exit Sub
              End If
              
               Call CloseRs(Rs)
               Call CloseDb(DbAnoAnt)
                    
                Q2 = "SELECT valor as Total FROM BaseImponible14D"
                Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 AND VALOR = " & Total
            
               Set Rs2 = OpenRs(DbMain, Q2)
            
               If Not Rs2.EOF Then
                 MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
               Else
               
               Call CloseRs(Rs2)
                         
                      Q2 = ""
                       
                      Q2 = "SELECT valor as Total FROM BaseImponible14D"
                      Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                      Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 "
            
                     Set Rs3 = OpenRs(DbMain, Q2)
                     If Not Rs3.EOF Then
                       MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
                     Else
                     
                        Q1 = "INSERT INTO BaseImponible14D (IdEmpresa, Ano, Tipo, Nivel, Codigo, Fecha, Valor)"
                        Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & lTipo
                        Q1 = Q1 & ", " & BIMP14D_MAXNIV + 1 & ", 9600, " & Fecha & ", " & vFmt(Total * -1) & ")"
                  
                        Call ExecSQL(DbMain, Q1)
                        
                        End If
                        
                        Call CloseRs(Rs3)
             End If
      Else
        Call CloseRs(Rs)
      End If
  
    
'Else
#Else
    If gEmpresa.TieneAnoAntAccess Then
       DbName = Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad", "")
       If ExistFile(DbName) Then
          'Call OpenDbEmp2(DbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(DbName, False, False, ConnStr)
          'Set DbAnoAnt = OpenConnection(DbName, False, False, ConnStr)
       End If
    End If
    
          Q1 = "SELECT iif(sum(valor) is null, 0,sum(valor)) as Total FROM BaseImponible14D"
          Q1 = Q1 & " WHERE ano = " & IniAnoAnterior
          Q1 = Q1 & " AND FECHA = 0 "
      
      If gEmpresa.TieneAnoAntAccess Then
        Set Rs5 = OpenRsDao(DbAnoAnt, Q1)
      Else
        Set Rs = OpenRs(DbMain, Q1)
      End If

    If Not Rs Is Nothing Then
        If Not Rs.EOF Then
        Total = vFld(Rs("Total"))
        
              If Total >= 0 Then
                  Call CloseRs(Rs)
                    If gEmpresa.TieneAnoAntAccess Then
                    Call CloseDb(DbAnoAnt)
                    End If
                MsgBox1 "No existe Perdida del Año Anterior", vbExclamation
                Exit Sub
              End If
              
                  Q2 = "SELECT valor as Total FROM BaseImponible14D"
                  Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                  Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 AND VALOR = " & Total * -1
                              
              
                 Set Rs2 = OpenRs(DbMain, Q2)
              
                 If Not Rs2.EOF Then
                   MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
                 Else
                 
                 Call CloseRs(Rs2)
                   
                          Q2 = ""
                           
                          Q2 = "SELECT valor as Total FROM BaseImponible14D"
                          Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                          Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 "
                
                         Set Rs3 = OpenRs(DbMain, Q2)
                         If Not Rs3.EOF Then
                           MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
                         Else
                         
                            Q1 = "INSERT INTO BaseImponible14D (IdEmpresa, Ano, Tipo, Nivel, Codigo, Fecha, Valor)"
                            Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & lTipo
                            Q1 = Q1 & ", " & BIMP14D_MAXNIV + 1 & ", 9600, " & Fecha & ", " & vFmt(Total * -1) & ")"
                      
                            Call ExecSQL(DbMain, Q1)
                            
                            End If
                            
                            Call CloseRs(Rs3)
                 End If
          Else
            Call CloseRs(Rs)
          End If
      Else
        If Not Rs5.EOF Then
        Total = Val(Rs5("Total"))
        
              If Total >= 0 Then
                  Call CloseRs(Rs5)
                    If gEmpresa.TieneAnoAntAccess Then
                    Call CloseDb(DbAnoAnt)
                    End If
                MsgBox1 "No existe Perdida del Año Anterior", vbExclamation
                Exit Sub
              End If
              
                  Q2 = "SELECT valor as Total FROM BaseImponible14D"
                  Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                  Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 AND VALOR = " & Total * -1
                              
              
                 Set Rs2 = OpenRs(DbMain, Q2)
              
                 If Not Rs2.EOF Then
                   MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
                 Else
                 
                 Call CloseRs(Rs2)
                   
                          Q2 = ""
                           
                          Q2 = "SELECT valor as Total FROM BaseImponible14D"
                          Q2 = Q2 & " WHERE ano = " & gEmpresa.Ano
                          Q2 = Q2 & " AND CODIGO = 9600 AND FECHA > 0 "
                
                         Set Rs3 = OpenRs(DbMain, Q2)
                         If Not Rs3.EOF Then
                           MsgBox1 "Ya se encuetra registrada Perdida del Año Anterior", vbExclamation
                         Else
                         
                            Q1 = "INSERT INTO BaseImponible14D (IdEmpresa, Ano, Tipo, Nivel, Codigo, Fecha, Valor)"
                            Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & lTipo
                            Q1 = Q1 & ", " & BIMP14D_MAXNIV + 1 & ", 9600, " & Fecha & ", " & vFmt(Total * -1) & ")"
                      
                            Call ExecSQL(DbMain, Q1)
                            
                            End If
                            
                            Call CloseRs(Rs3)
                 End If
          Else
            Call CloseRs(Rs5)
          End If
          
          
      End If
    
 'End If
 #End If
             
 'Call CloseRs(Rs2)
 'Call CloseRs(Rs3)
End Sub

'fin 2699582



