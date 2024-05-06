Attribute VB_Name = "Mod14D"
Option Explicit
'Valida una entidad relacionada para 14 D
Public Function ValidaEnt14D(ByVal IdEnt As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   If IdEnt = 0 Then
      ValidaEnt14D = False
      Exit Function
   End If
   
   If gEmpresa.Ano < 2020 Then
      ValidaEnt14D = True
      Exit Function
   End If
   
   ValidaEnt14D = False
   
   If gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp Then
      Q1 = "SELECT EntRelacionada, FranqTribEnt FROM Entidades WHERE IdEntidad = " & IdEnt
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         If vFld(Rs("EntRelacionada")) <> 0 And vFld(Rs("FranqTribEnt")) = FTE_14A Then
            ValidaEnt14D = True
         End If
      End If
      Call CloseRs(Rs)
   End If
   
End Function

'retorna código de ajuste para 14A, 14D Nro. 3 o 14D Nro. 8, dependiendo de regimen al que está acogido el cliente y el año
Public Function HomologaCod14D(ByVal CodAjuste As Integer) As Integer

   If gEmpresa.Ano < 2020 Then
      HomologaCod14D = CodAjuste
      Exit Function
   End If
   
   HomologaCod14D = CodAjuste
   
   If CodAjuste < 1400 Then   'es código Antiguo, anterior a 2020
   
      If gEmpresa.ProPymeGeneral Then
      
         Select Case CodAjuste
         
            Case 628
               HomologaCod14D = 1400
            Case 651
               HomologaCod14D = 1588
            Case 630
               HomologaCod14D = 1409
            Case 631
               HomologaCod14D = 1411
            Case 632
               HomologaCod14D = 1413
            Case 633
               HomologaCod14D = 1419
            Case 635
               HomologaCod14D = 1424
            Case 1140 'pipe enero 2022 tema 4 2738156
               HomologaCod14D = 1415
               'fin
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
         
      ElseIf gEmpresa.ProPymeTransp Then
               
         Select Case CodAjuste
         
            Case 628
               HomologaCod14D = 1600
            Case 651
               HomologaCod14D = 1607
            Case 630
               HomologaCod14D = 1614
            Case 631
               HomologaCod14D = 1616
            Case 632
               HomologaCod14D = 1618
            Case 633
               HomologaCod14D = 1622
            Case 635
               HomologaCod14D = 1625
            Case 1140 'pipe enero 2022 tema 4 2738156
               HomologaCod14D = 1620
               'fin
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
         
      ElseIf gEmpresa.R14ASemiIntegrado Then
               
         Select Case CodAjuste
         
            Case 628
               HomologaCod14D = 1657
            Case 851
               HomologaCod14D = 1658
            Case 629
               HomologaCod14D = 1659
            Case 651
               HomologaCod14D = 1660
            Case 630
               HomologaCod14D = 1661
            Case 631
               HomologaCod14D = 1662
            Case 632
               HomologaCod14D = 1663
            Case 633
               HomologaCod14D = 1664
            Case 635
               HomologaCod14D = 1671
            Case 967
               HomologaCod14D = 1666
            Case 852
               HomologaCod14D = 1667
            Case 897
               HomologaCod14D = 1668
            Case 853
               HomologaCod14D = 1669
            Case 968
               HomologaCod14D = 1670
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
         
      End If
   
   Else        'es código 14D o 14A
            
      If gEmpresa.ProPymeGeneral Then
       
         Select Case CodAjuste
         
            Case 1400, 1600, 1657         '628
               HomologaCod14D = 1400
            Case 1588, 1607, 1660         '651
               HomologaCod14D = 1588
            Case 1409, 1614, 1661         '630
               HomologaCod14D = 1409
            Case 1411, 1616, 1662         '631
               HomologaCod14D = 1411
            Case 1413, 1618, 1663         '632
               HomologaCod14D = 1413
            Case 1419, 1622, 1664         '633
               HomologaCod14D = 1419
            Case 1424, 1625, 1671         '635
               HomologaCod14D = 1424
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
      
      ElseIf gEmpresa.ProPymeTransp Then
       
         Select Case CodAjuste
         
            Case 1400, 1600, 1657         '628
               HomologaCod14D = 1600
            Case 1588, 1607, 1660         '651
               HomologaCod14D = 1607
            Case 1409, 1614, 1661         '630
               HomologaCod14D = 1614
            Case 1411, 1616, 1662         '631
               HomologaCod14D = 1616
            Case 1413, 1618, 1663         '632
               HomologaCod14D = 1618
            Case 1419, 1622, 1664         '633
               HomologaCod14D = 1622
            Case 1424, 1625, 1671         '635
               HomologaCod14D = 1625
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
         
      ElseIf gEmpresa.R14ASemiIntegrado Then
      
         Select Case CodAjuste
         
            Case 1400, 1600, 1657         '628
               HomologaCod14D = 1600
            Case 1588, 1607, 1660         '651
               HomologaCod14D = 1607
            Case 1409, 1614, 1661         '630
               HomologaCod14D = 1614
            Case 1411, 1616, 1662         '631
               HomologaCod14D = 1616
            Case 1413, 1618, 1663         '632
               HomologaCod14D = 1618
            Case 1419, 1622, 1664         '633
               HomologaCod14D = 1622
            Case 1424, 1625, 1671         '635
               HomologaCod14D = 1671
            Case Else
               HomologaCod14D = CodAjuste
            
         End Select
       
      Else
         Select Case CodAjuste
            Case 1400, 1600, 1657         '628
               HomologaCod14D = 628
            Case 1588, 1607, 1660         '651
               HomologaCod14D = 651
            Case 1409, 1614, 1661         '630
               HomologaCod14D = 630
            Case 1411, 1616, 1662         '631
               HomologaCod14D = 631
            Case 1413, 1618, 1663         '632
               HomologaCod14D = 632
            Case 1419, 1622, 1664         '633
               HomologaCod14D = 633
            Case 1424, 1625, 1671         '635
               HomologaCod14D = 635
            Case 1658                     '851
               HomologaCod14D = 851
            Case 1659                     '629
               HomologaCod14D = 629
            Case 1665                     '966
               HomologaCod14D = 966
            Case 1666                     '967
               HomologaCod14D = 967
            Case 1667                     '852
               HomologaCod14D = 852
            Case 1668                     '897
               HomologaCod14D = 897
            Case 1669                     '853
               HomologaCod14D = 853
            Case 1670                     '968
               HomologaCod14D = 968
           Case Else
               HomologaCod14D = CodAjuste
               
         End Select
         
      End If
   
   End If

End Function

'recorre todo el plan de cuentas y homologa el Cod_F2214Ter a 14D si corresponde
Public Function Homologa_CodF22_14Ter_14D()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim CodF22 As Integer
   
   Q1 = "SELECT IdCuenta, CodF22_14Ter FROM Cuentas WHERE CodF22_14Ter <> 0"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      CodF22 = HomologaCod14D(vFld(Rs("CodF22_14Ter")))
   
      If vFld(Rs("CodF22_14Ter")) <> CodF22 Then
         Q1 = "UPDATE Cuentas SET CodF22_14Ter = " & CodF22
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
End Function
