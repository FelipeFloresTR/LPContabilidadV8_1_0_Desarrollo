Attribute VB_Name = "TraspasosHR"
Option Explicit

'agregar a SourceSafe

'Para links con DB HR
Public gLinkF22 As Boolean       ' Se pudo linkear las tablas del Form 22
Public gLinkParFUT As Boolean    ' Se pudo linkear la tabla de Parámetros de FUT
Public gPathForm22 As String     'path Form 22 = "\FORM22"
Public gPathPlan22 As String     'path Planificación Form 22 = "\PLAN22"


'Constantes HR_LAU de Tipos de Documentos
Public Const LAU_COMPRASNAC = 30       'Compras nacionales
Public Const LAU_COMPRASIMP = 31       'Compras Importadas
Public Const LAU_VENTASBOLETA = 42     'Ventas con Boleta
Public Const LAU_VENTASDEVBOL = 43     'Devolución Ventas con Boleta
Public Const LAU_VENTASEXP = 41        'Ventas exportaciones
Public Const LAU_VENTASNAC = 40        'Ventas Nacionales

'Libros HR-LAU
Public Const LAU_LIBCOMPRAS = 3
Public Const LAU_LIBVENTAS = 4
Public Const LAU_LIBRETEN = 2
Public Const LAU_LIBREMU = 1

'Tipo Docs LAU Libro de Compras
Public Const LAU_COMP_FACT = 0         'Factura
Public Const LAU_COMP_NOTADEB = 1      'Nota de débito
Public Const LAU_COMP_NOTACRED = 2     'Nota de Crédito
Public Const LAU_COMP_FACTCOMP = 3     'Factura de compra
Public Const LAU_COMP_OTRO = 4         'Otro
Public Const LAU_COMP_FACTEXEN = 5     'Factura Exenta
Public Const LAU_COMP_FACTIMP = 6      'Factura de Importación

'Tipo Docs LAU Libro de Ventas
Public Const LAU_VENTA_FACT = 0        'Factura
Public Const LAU_VENTA_NOTADEB = 1     'Nota de débito
Public Const LAU_VENTA_NOTACRED = 2    'Nota de Crédito
Public Const LAU_VENTA_FACTEXEN = 3    'Factura Exenta
Public Const LAU_VENTA_OTRO = 4        'Otro
Public Const LAU_VENTA_FACTCOMP = 5    'Factura de compra

Public Const LAU_VENTA_LIQFACT = 6     'Liquidación Factura
Public Const LAU_VENTA_BOLETA = 7      'venta con boletas
Public Const LAU_VENTA_DEVBOLETA = 8   'venta devolución con boleta
Public Const LAU_VENTA_FACTEXP = 9     'Factura exportación
Public Const LAU_VENTA_NCREDEXP = 10   'Nota Crédito de exportación
Public Const LAU_VENTA_NDEBEXP = 11    'Nota Débito de exportación
Public Const LAU_VENTA_BOLEXENTA = 12  'venta con boletas exentas
Public Const LAU_VENTA_VTAMENOR = 13   'ventas menores
Public Const LAU_VENTA_VALEPAGOELECTR = 14   'VPE: vale pago electrónico
'2814014 pipe
Public Const LAU_VENTA_BOLVENTAEXENTA = 15  'VPEE Vale Pago Electronico con Exento
'fin 2814014


'Tipos de Ingresos o Gastos de FUT
Public Const FUT_AGRPAG = 1            'Agregados Pagados
Public Const FUT_AGRADE = 2            'Agregados Adeudados
Public Const FUT_DEDPER = 3            'Deducciones Percibidas
Public Const FUT_DEDDEV = 4            'Deducciones Devengadas
Public Const FUT_AMBOS = 5

Public gTipoIngGasFUT(FUT_DEDDEV) As String

'Tipo de contribuyente para FUT (YA NO SE USAN, SE USA CONTRIB_)
'Public Const FUTCONT_EMPIND = 1        'Empresario Individual
'Public Const FUTCONT_OTSOC = 2         'otras sociedades
'Public Const FUTCONT_SACER = 3         'S.A. Cerrada
'Public Const FUTCONT_SAABI = 4         'S.A. Abierta


Type CuentaFUT_t
   id As Long                          'correlativo en tabla CuentasFUT
   TipoIngGas As Integer              'tipo de ingreso o gasto
   IdItemFUT As Integer                'IdItem en tabla FUT
   IdCuenta As Long
   CodCuenta As String
   Descrip As String
End Type

#If DATACON = 1 Then
Public Function IniTraspasos()

   gTipoIngGasFUT(FUT_AGRPAG) = "Agregados Pagados"
   gTipoIngGasFUT(FUT_AGRADE) = "Agregados Adeudados"
   gTipoIngGasFUT(FUT_DEDPER) = "Deducciones Percibidas"
   gTipoIngGasFUT(FUT_DEDDEV) = "Deducciones Devengadas"
   
End Function

Public Sub GetIngGas(RsRes As Recordset, ConDetalle As Boolean, ByVal TipoIngGas As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim LCod As Integer
   Dim Tbl As TableDef
   Dim DetStr As String
   Dim DetJoin As String
   Dim WhTipo As String
   Dim DetOrderBy As String

   On Error Resume Next
   Call ExecSQL(DbMain, "DROP TABLE TmpExpFUT001")
   On Error GoTo 0

   'creamos la tabla temporal
   Q1 = "SELECT 0 as TipoIngGas, 0 as IdItem, 0 as IdCuenta, '" & String(15, "0") & "' As CodCuenta, '" & String(100, " ") & "' As DescCuenta INTO TmpExpFUT001"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "SELECT TipoIngGas, IdItem, CuentasFUT.IdCuenta, Codigo, Nivel, Descripcion "
   Q1 = Q1 & " FROM CuentasFUT INNER JOIN Cuentas ON CuentasFUT.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & " ORDER BY TipoIngGas, IdItem "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      If vFld(Rs("Nivel")) <> gLastNivel Then
         LCod = gNiveles.Inicio(vFld(Rs("Nivel")) + 1) - 1
   
         Q1 = "INSERT INTO TmpExpFUT001 "
         Q1 = Q1 & " SELECT " & vFld(Rs("TipoIngGas")) & " As TipoIngGas, "
         Q1 = Q1 & vFld(Rs("IdItem")) & " As IdItem, IdCuenta as IdCuenta, Cuentas.Codigo as CodCuenta, Cuentas.Descripcion as DescCuenta "
         Q1 = Q1 & " FROM Cuentas "
         Q1 = Q1 & " WHERE left(Codigo," & LCod & ") = '" & Left(vFld(Rs("Codigo")), LCod) & "'"
         Q1 = Q1 & " and Nivel = " & gLastNivel
   
      Else
         Q1 = "INSERT INTO TmpExpFUT001 (TipoIngGas, IdItem, IdCuenta, CodCuenta, DescCuenta ) "
         Q1 = Q1 & "VALUES(" & vFld(Rs("TipoIngGas")) & "," & vFld(Rs("IdItem")) & "," & vFld(Rs("IdCuenta")) & ",'" & vFld(Rs("Codigo")) & "','" & vFld(Rs("Descripcion")) & "')"
      
      End If
      
      Call ExecSQL(DbMain, Q1)
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If ConDetalle Then
      'DetStr = ", TmpExpFUT001.IdItem, HR_FUTGrItems.GrpOIte, HR_FUTGrItems.Descripci, Comprobante.Fecha, MovComprobante.Glosa, COrden "
      DetStr = ", TmpExpFUT001.IdItem, HR_FUTGrItems.GrpOIte, HR_FUTGrItems.Descripci, Comprobante.Fecha, TmpExpFUT001.DescCuenta, COrden "
      DetJoin = " INNER JOIN HR_FUTGrItems ON Val(HR_FUTGrItems.IdItem) = TmpExpFUT001.IdItem "
      DetOrderBy = ", COrden, Comprobante.Fecha "
   End If
   
   If TipoIngGas > 0 Then
      WhTipo = "WHERE TipoIngGas = " & TipoIngGas
   End If
   
   Q1 = "SELECT TmpExpFUT001.TipoIngGas " & DetStr & ","
   Q1 = Q1 & " Sum(abs(MovComprobante.Debe - MovComprobante.Haber)) as Valor "
   Q1 = Q1 & " FROM ((MovComprobante INNER JOIN TmpExpFUT001 ON MovComprobante.IdCuenta = TmpExpFUT001.IdCuenta)"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp) "
   Q1 = Q1 & DetJoin
   Q1 = Q1 & WhTipo
   Q1 = Q1 & " GROUP BY TmpExpFUT001.TipoIngGas " & DetStr
   Q1 = Q1 & " ORDER BY TmpExpFUT001.TipoIngGas " & DetOrderBy

   Set RsRes = OpenRs(DbMain, Q1)
   
   Call ExecSQL(DbMain, "DROP TABLE TmpExpFUT001")
  
End Sub
#End If
