Attribute VB_Name = "ModCPS"
Option Explicit

'Obtiene el valor anual del CPS para un determinado registro del CPS, sumando el detalle
Public Function GetCPSAnual(ByVal TipoDetCPS As Integer) As Double
   Dim Q1 As String
   Dim Rs As Recordset

   GetCPSAnual = 0
   Q1 = "SELECT Valor FROM CapPropioSimplAnual "
   Q1 = Q1 & " WHERE TipoDetCPS = " & TipoDetCPS & " AND AnoValor = " & gEmpresa.Ano
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id

   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      GetCPSAnual = vFld(Rs("Valor"))
   Else
      GetCPSAnual = 0
   End If
   
   Call CloseRs(Rs)
   
End Function
