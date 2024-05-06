Attribute VB_Name = "ModExpSII"
Option Explicit

Public Const MAX_CODFLDSII_VENTAS = 24
Public Const MAX_CODFLDSII_COMPRAS = 28


Public gFldSIIVentas(MAX_CODFLDSII_VENTAS) As String
Public gFldSIICompras(MAX_CODFLDSII_COMPRAS) As String

Type ImpAdicDoc_t
   IdTipoValLib As Long
   CodSIIDTE As String
   Valor As Double
   Tasa As Single
End Type

Public Const MAX_IMPADICDOC = 20

Public Function IniFldExpSII()

   Call ClearFLDSIIVentas
   Call ClearFLDSIICompras
   
End Function

Private Function ClearFLDSIIVentas()
   
   gFldSIIVentas(1) = FillStrR("1", 9)    'Rut Contrubuyente  con dígito sin guión
   gFldSIIVentas(2) = FillStrR("2", 6)    'Fecha de Registro  MMAAAA
   gFldSIIVentas(3) = FillStrR("3", 10)   'Número interno
   gFldSIIVentas(4) = FillStrR("4", 1)    'Tipo De operación "V" Ventas
   gFldSIIVentas(5) = FillStrR("5", 3)    'Tipo de documento
   gFldSIIVentas(6) = FillStrR("6", 10)   'Número de documento
   gFldSIIVentas(7) = FillStrR("7", 8)    'Fecha de documento DDMMAAAA
   gFldSIIVentas(8) = FillStrR("8", 9)    'Rut Asociado (Cliente) con dígito sin guióm
   gFldSIIVentas(9) = FillStrR("9", 50)   'Nombre o razón social
   gFldSIIVentas(10) = FillStrR("10", 13) 'monto exento
   gFldSIIVentas(11) = FillStrR("11", 13) 'monto neto
   gFldSIIVentas(12) = FillStrR("12", 13) 'monto IVA
   gFldSIIVentas(13) = FillStrR("13", 13) 'monto total operación
   
   gFldSIIVentas(14) = FillStrR("39", 13)    'IVA Rentenido Total
   gFldSIIVentas(15) = FillStrR("42", 13)    'IVA Rentenido Parcial
   gFldSIIVentas(16) = FillStrR("126", 13)   'Crédito especial 65% enpresas constructoras
   gFldSIIVentas(17) = FillStrR("127", 13)   'Impuesto específico a los combustibles
   gFldSIIVentas(18) = FillStrR("113", 13)   'Impuesto adicional a cierttos productos
   gFldSIIVentas(19) = FillStrR("148", 13)   'Impuesto Art.42, Letra a
   gFldSIIVentas(20) = FillStrR("45", 13)   'Impuesto Art.42, Letra b
   gFldSIIVentas(21) = FillStrR("32", 13)   'Impuesto Art.42, Letra c
   gFldSIIVentas(22) = FillStrR("150", 13)   'Impuesto Art.42, Letra c
   gFldSIIVentas(23) = FillStrR("146", 13)   'Impuesto Art.42, Letra d y e
   gFldSIIVentas(24) = FillStrR("31", 13)   'Impuesto Art.42, Letra f
   

End Function

Private Function ClearFLDSIICompras()
   
   gFldSIICompras(1) = FillStrR("1", 9)    'Rut Contrubuyente  con dígito sin guión
   gFldSIICompras(2) = FillStrR("2", 6)    'Fecha de Registro  MMAAAA
   gFldSIICompras(3) = FillStrR("3", 10)   'Número interno
   gFldSIICompras(4) = FillStrR("4", 1)    'Tipo De operación "C" Compra
   gFldSIICompras(5) = FillStrR("5", 3)    'Tipo de documento
   gFldSIICompras(6) = FillStrR("6", 10)   'Número de documento
   gFldSIICompras(7) = FillStrR("7", 8)    'Fecha de documento DDMMAAAA
   gFldSIICompras(8) = FillStrR("8", 9)    'Rut Asociado (Proveedor) con dígito sin guióm
   gFldSIICompras(9) = FillStrR("9", 50)   'Nombre o razón social
   gFldSIICompras(10) = FillStrR("10", 13) 'monto exento
   gFldSIICompras(11) = FillStrR("11", 13) 'monto neto
   gFldSIICompras(12) = FillStrR("12", 13) 'monto IVA
   gFldSIICompras(13) = FillStrR("13", 13) 'monto total operación
   
   gFldSIICompras(14) = FillStrR("14", 13)    'IVA No Recuperable
   gFldSIICompras(15) = FillStrR("15", 13)    'IVA Uso Común
   gFldSIICompras(16) = FillStrR("39", 13)    'IVA Retenido Total
   gFldSIICompras(17) = FillStrR("42", 13)    'IVA Retenido Parcial
   gFldSIICompras(18) = FillStrR("126", 13)   'Crédito especial 65% enpresas constructoras
   gFldSIICompras(19) = FillStrR("127", 13)   'Impuesto específico a los combustibles
   gFldSIICompras(20) = FillStrR("28", 13)    'Impuesto adicional a cierttos productos
   gFldSIICompras(21) = FillStrR("147", 13)   'Impuesto Art.42, Letra a
   gFldSIICompras(22) = FillStrR("27", 13)    'Impuesto Art.42, Letra b
   gFldSIICompras(23) = FillStrR("33", 13)    'Impuesto Art.42, Letra c
   gFldSIICompras(24) = FillStrR("149", 13)   'Impuesto Art.42, Letra c
   gFldSIICompras(25) = FillStrR("85", 13)    'Impuesto Art.42, Letra d y e
   gFldSIICompras(26) = FillStrR("87", 13)    'Impuesto Art.42, Letra f
   gFldSIICompras(27) = FillStrR("160", 13)   'IVA anticipado
   gFldSIICompras(28) = FillStrR("180", 13)   'Activo fijo
   

End Function


Public Function GenRegLargoFijo(FldSII() As String, FldArray() As String) As String
   Dim Reg As String
   Dim i As Integer
   
   For i = 1 To UBound(FldSII)
      Reg = Reg & FillStrR(FldArray(i), Len(FldSII(i)))
   Next i
   
   GenRegLargoFijo = Reg
   
End Function

