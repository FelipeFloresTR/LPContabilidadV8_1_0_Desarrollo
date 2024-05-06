Attribute VB_Name = "LpContab_ANSi_Textos"
Option Explicit

'Esta función se desarrolló para corregir el error que generó el primer Script de la base de datos
'que quedó grabado en UTF8 en vez de ANSI
Function CorrigeTextosConAcentosScriptUTF8()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Txt As String, Txt2 As String, Txt3 As String, Txt4 As String
   Dim bUpdate As Boolean
   
   'CodActiv
   Q1 = "SELECT Codigo, Descrip, Version FROM CodActiv"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descrip"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CodActiv SET Descrip = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Codigo = '" & vFld(Rs("Codigo")) & "' AND Version = " & vFld(Rs("Version"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_ComprobanteBase
   Q1 = "SELECT IdComp, Nombre, Descrip, Glosa FROM CT_ComprobanteBase "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("Descrip"))
      Txt3 = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE CT_ComprobanteBase SET Nombre = '" & ParaSQL(Txt) & "', Descrip = '" & ParaSQL(Txt2) & "', Glosa = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'CT_MovComprobanteBase
   Q1 = "SELECT IdMov, IdComp, Glosa FROM CT_MovComprobanteBase"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CT_MovComprobanteBase SET Glosa = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdMov = '" & vFld(Rs("IdMov")) & "' AND IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_Comprobante
   Q1 = "SELECT IdComp, Nombre, Descrip, Glosa FROM CT_Comprobante "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("Descrip"))
      Txt3 = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE CT_Comprobante SET Nombre = '" & ParaSQL(Txt) & "', Descrip = '" & ParaSQL(Txt2) & "', Glosa = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'CT_MovComprobante
   Q1 = "SELECT IdMov, IdComp, Glosa FROM CT_MovComprobante"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Glosa"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE CT_MovComprobante SET Glosa = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdMov = '" & vFld(Rs("IdMov")) & "' AND IdComp = " & vFld(Rs("IdComp"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'IFRS_PlanIFRS
   Q1 = "SELECT IdCuenta, Descripcion FROM IFRS_PlanIFRS"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE IFRS_PlanIFRS SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'PlanBasico
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanBasico"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanBasico SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'PlanIntermedio
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanIntermedio"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanIntermedio SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)

   'PlanAvanzado
   Q1 = "SELECT IdCuenta, Descripcion FROM PlanAvanzado"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanAvanzado SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'Cuentas
   Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Descripcion"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE Cuentas SET Descripcion = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdCuenta = " & vFld(Rs("IdCuenta"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'PlanCuentasSII
   Q1 = "SELECT IdPlanCuentasSII, DescripSII FROM PlanCuentasSII"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("DescripSII"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE PlanCuentasSII SET DescripSII = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE IdPlanCuentasSII = " & vFld(Rs("IdPlanCuentasSII"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'RazonesFin
   Q1 = "SELECT IdRazon, Nombre, TxtNumerador, TxtDenominador FROM RazonesFin "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      Txt2 = vFld(Rs("TxtNumerador"))
      Txt3 = vFld(Rs("TxtDenominador"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE RazonesFin SET Nombre = '" & ParaSQL(Txt) & "', TxtNumerador = '" & ParaSQL(Txt2) & "', TxtDenominador = '" & ParaSQL(Txt3) & "'"
         Q1 = Q1 & " WHERE IdRazon = " & vFld(Rs("IdRazon"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   
   'Regiones
   Q1 = "SELECT Id, Comuna FROM Regiones"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Comuna"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE Regiones SET Comuna = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Id = " & vFld(Rs("Id"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'TipoDocs
   Q1 = "SELECT Id, Nombre FROM TipoDocs"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Nombre"))
      
      bUpdate = CorrigeTexto(Txt)
      If bUpdate Then
         Q1 = "UPDATE TipoDocs SET Nombre = '" & ParaSQL(Txt) & "'"
         Q1 = Q1 & " WHERE Id = " & vFld(Rs("Id"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   'TipoValor
   Q1 = "SELECT IdTValor, Valor, Tit1, Tit2, TitCompleto FROM TipoValor "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Txt = vFld(Rs("Valor"))
      Txt2 = vFld(Rs("Tit1"))
      Txt3 = vFld(Rs("Tit2"))
      Txt4 = vFld(Rs("TitCompleto"))
      
      bUpdate = CorrigeTexto(Txt) Or CorrigeTexto(Txt2) Or CorrigeTexto(Txt3)
      If bUpdate Then
         Q1 = "UPDATE TipoValor SET Valor = '" & ParaSQL(Txt) & "', Tit1 = '" & ParaSQL(Txt2) & "', Tit2 = '" & ParaSQL(Txt3) & "', TitCompleto = '" & ParaSQL(Txt4) & "'"
         Q1 = Q1 & " WHERE IdTValor = " & vFld(Rs("IdTValor"))
         
      End If
      
      Rs.MoveNext
      
      If bUpdate Then
         Call ExecSQL(DbMain, Q1)
      End If
   
   Loop
   
   Call CloseRs(Rs)
   

   MsgBox1 "Proceso finalizado", vbOKOnly + vbInformation
   
End Function



Function CorrigeTexto(Buf As String) As Boolean
    If InStr(Buf, "?") <= 0 Then
       Exit Function
    End If

   Buf = Replace(Buf, "ma?z", "maíz")
   Buf = Replace(Buf, "tub?rculos", "tubérculos")
   Buf = Replace(Buf, "ca?a", "caña")
   Buf = Replace(Buf, "az?car", "azúcar")
   Buf = Replace(Buf, "producci?n", "producción")
   Buf = Replace(Buf, "c?tricos", "cítricos")
   Buf = Replace(Buf, "?rboles", "árboles")
   Buf = Replace(Buf, "caf?", "café")
   Buf = Replace(Buf, "t?", "té")
   Buf = Replace(Buf, "arom?ticas", "aromáticas")
   Buf = Replace(Buf, "farmac?uticas", "farmacéuticas")
   Buf = Replace(Buf, "Cr?a", "Cría")
   Buf = Replace(Buf, "vicu?as", "vicuñas")
   Buf = Replace(Buf, "cam?lidos", "camélidos")
   Buf = Replace(Buf, "agr?colas", "agrícolas")
   Buf = Replace(Buf, "combinaci?n", "combinación")
   Buf = Replace(Buf, "cr?a", "cría")
   Buf = Replace(Buf, "explotaci?n", "explotación")
   Buf = Replace(Buf, "ganader?a", "ganadería")
   Buf = Replace(Buf, "propagaci?n", "propagación")
   Buf = Replace(Buf, "Explotaci?n", "Explotación")
   Buf = Replace(Buf, "Extracci?n", "Extracción")
   Buf = Replace(Buf, "Recolecci?n", "Recolección")
   Buf = Replace(Buf, "forestaci?n", "forestación")
   Buf = Replace(Buf, "retribuci?n", "retribución")
   Buf = Replace(Buf, "extinci?n", "extinción")
   Buf = Replace(Buf, "prevenci?n", "prevención")
   Buf = Replace(Buf, "mar?tima", "marítima")
   Buf = Replace(Buf, "factor?a", "factoría")
   Buf = Replace(Buf, "extracci?n", "extracción")
   Buf = Replace(Buf, "reproducci?n", "reproducción")
   Buf = Replace(Buf, "Reproducci?n", "Reproducción")
   Buf = Replace(Buf, "crust?ceos", "crustáceos")
   Buf = Replace(Buf, "carb?n", "carbón")
   Buf = Replace(Buf, "petr?leo", "petróleo")
   Buf = Replace(Buf, "metal?feros", "metalíferos")
   Buf = Replace(Buf, "fabricaci?n", "fabricación")
   Buf = Replace(Buf, "qu?micos", "químicos")
   Buf = Replace(Buf, "Elaboraci?n", "Elaboración")
   Buf = Replace(Buf, "conservaci?n", "conservación")
   Buf = Replace(Buf, "c?rnicos", "cárnicos")
   Buf = Replace(Buf, "Producci?n", "Producción")
   Buf = Replace(Buf, "salm?nidos", "salmónidos")
   Buf = Replace(Buf, "acu?ticos", "acuáticos")
   Buf = Replace(Buf, "elaboraci?n", "elaboración")
   Buf = Replace(Buf, "l?cteos", "lácteos")
   Buf = Replace(Buf, "s?mola", "sémola")
   Buf = Replace(Buf, "gr?nulos", "gránulos")
   Buf = Replace(Buf, "moliner?a", "molinería")
   Buf = Replace(Buf, "almid?n", "almidón")
   Buf = Replace(Buf, "panader?a", "panadería")
   Buf = Replace(Buf, "pasteler?a", "pastelería")
   Buf = Replace(Buf, "confiter?a", "confitería")
   Buf = Replace(Buf, "farin?ceos", "farináceos")
   Buf = Replace(Buf, "informaci?n", "información")
   Buf = Replace(Buf, "Destilaci?n", "Destilación")
   Buf = Replace(Buf, "rectificaci?n", "rectificación")
   Buf = Replace(Buf, "alcoh?licas", "alcohólicas")
   Buf = Replace(Buf, "Preparaci?n", "Preparación")
   Buf = Replace(Buf, "Fabricaci?n", "Fabricación")
   Buf = Replace(Buf, "art?culos", "artículos")
   Buf = Replace(Buf, "te?ido", "teñido")
   Buf = Replace(Buf, "talabarter?a", "talabartería")
   Buf = Replace(Buf, "guarnicioner?a", "guarnicionería")
   Buf = Replace(Buf, "carpinter?a", "carpintería")
   Buf = Replace(Buf, "cart?n", "cartón")
   Buf = Replace(Buf, "Impresi?n", "Impresión")
   Buf = Replace(Buf, "impresi?n", "impresión")
   Buf = Replace(Buf, "refinaci?n", "refinación")
   Buf = Replace(Buf, "qu?micas", "químicas")
   Buf = Replace(Buf, "b?sicas", "básicas")
   Buf = Replace(Buf, "nitr?geno", "nitrógeno")
   Buf = Replace(Buf, "pl?sticos", "plásticos")
   Buf = Replace(Buf, "sint?tico", "sintético")
   Buf = Replace(Buf, "pirot?cnicos", "pirotécnicos")
   Buf = Replace(Buf, "farmac?uticos", "farmacéuticos")
   Buf = Replace(Buf, "bot?nicos", "botánicos")
   Buf = Replace(Buf, "c?maras", "cámaras")
   Buf = Replace(Buf, "renovaci?n", "renovación")
   Buf = Replace(Buf, "pl?stico", "plástico")
   Buf = Replace(Buf, "construcci?n", "construcción")
   Buf = Replace(Buf, "cer?mica", "cerámica")
   Buf = Replace(Buf, "hormig?n", "hormigón")
   Buf = Replace(Buf, "met?licos", "metálicos")
   Buf = Replace(Buf, "Fundici?n", "Fundición")
   Buf = Replace(Buf, "dep?sitos", "depósitos")
   Buf = Replace(Buf, "calefacci?n", "calefacción")
   Buf = Replace(Buf, "cuchiller?a", "cuchillería")
   Buf = Replace(Buf, "ferreter?a", "ferretería")
   Buf = Replace(Buf, "electr?nicos", "electrónicos")
   Buf = Replace(Buf, "perif?rico", "periférico")
   Buf = Replace(Buf, "medici?n", "medición")
   Buf = Replace(Buf, "navegaci?n", "navegación")
   Buf = Replace(Buf, "irradiaci?n", "irradiación")
   Buf = Replace(Buf, "electr?nico", "electrónico")
   Buf = Replace(Buf, "m?dico", "médico")
   Buf = Replace(Buf, "terap?utico", "terapéutico")
   Buf = Replace(Buf, "?pticos", "ópticos")
   Buf = Replace(Buf, "fotogr?fico", "fotográfico")
   Buf = Replace(Buf, "magn?ticos", "magnéticos")
   Buf = Replace(Buf, "el?ctricos", "eléctricos")
   Buf = Replace(Buf, "distribuci?n", "distribución")
   Buf = Replace(Buf, "bater?as", "baterías")
   Buf = Replace(Buf, "?ptica", "óptica")
   Buf = Replace(Buf, "el?ctrico", "eléctrico")
   Buf = Replace(Buf, "iluminaci?n", "iluminación")
   Buf = Replace(Buf, "dom?stico", "doméstico")
   Buf = Replace(Buf, "veh?culos", "vehículos")
   Buf = Replace(Buf, "propulsi?n", "propulsión")
   Buf = Replace(Buf, "v?lvulas", "válvulas")
   Buf = Replace(Buf, "transmisi?n", "transmisión")
   Buf = Replace(Buf, "elevaci?n", "elevación")
   Buf = Replace(Buf, "manipulaci?n", "manipulación")
   Buf = Replace(Buf, "conformaci?n", "conformación")
   Buf = Replace(Buf, "m?quinas", "máquinas")
   Buf = Replace(Buf, "metal?rgica", "metalúrgica")
   Buf = Replace(Buf, "carrocer?as", "carrocerías")
   Buf = Replace(Buf, "Construcci?n", "Construcción")
   Buf = Replace(Buf, "bisuter?a", "bisutería")
   Buf = Replace(Buf, "m?dicos", "médicos")
   Buf = Replace(Buf, "oftalmol?gicos", "oftalmológicos")
   Buf = Replace(Buf, "odontol?gicos", "odontológicos")
   Buf = Replace(Buf, "Reparaci?n", "Reparación")
   Buf = Replace(Buf, "miner?a", "minería")
   Buf = Replace(Buf, "reparaci?n", "reparación")
   Buf = Replace(Buf, "dom?sticos", "domésticos")
   Buf = Replace(Buf, "Instalaci?n", "Instalación")
   Buf = Replace(Buf, "Generaci?n", "Generación")
   Buf = Replace(Buf, "energ?a", "energía")
   Buf = Replace(Buf, "el?ctrica", "eléctrica")
   Buf = Replace(Buf, "hidroel?ctricas", "hidroeléctricas")
   Buf = Replace(Buf, "termoel?ctricas", "termoeléctricas")
   Buf = Replace(Buf, "Transmisi?n", "Transmisión")
   Buf = Replace(Buf, "Distribuci?n", "Distribución")
   Buf = Replace(Buf, "Regasificaci?n", "Regasificación")
   Buf = Replace(Buf, "tuber?a", "tubería")
   Buf = Replace(Buf, "regasificaci?n", "regasificación")
   Buf = Replace(Buf, "Captaci?n", "Captación")
   Buf = Replace(Buf, "Evacuaci?n", "Evacuación")
   Buf = Replace(Buf, "eliminaci?n", "eliminación")
   Buf = Replace(Buf, "Recuperaci?n", "Recuperación")
   Buf = Replace(Buf, "descontaminaci?n", "descontaminación")
   Buf = Replace(Buf, "gesti?n", "gestión")
   Buf = Replace(Buf, "l?neas", "líneas")
   Buf = Replace(Buf, "p?blico", "público")
   Buf = Replace(Buf, "ingenier?a", "ingeniería")
   Buf = Replace(Buf, "Demolici?n", "Demolición")
   Buf = Replace(Buf, "el?ctricas", "eléctricas")
   Buf = Replace(Buf, "gasfiter?a", "gasfitería")
   Buf = Replace(Buf, "Terminaci?n", "Terminación")
   Buf = Replace(Buf, "perfumer?a", "perfumería")
   Buf = Replace(Buf, "cosm?ticos", "cosméticos")
   Buf = Replace(Buf, "papeler?a", "papelería")
   Buf = Replace(Buf, "cient?ficos", "científicos")
   Buf = Replace(Buf, "quir?rgicos", "quirúrgicos")
   Buf = Replace(Buf, "inform?ticos", "informáticos")
   Buf = Replace(Buf, "s?lidos", "sólidos")
   Buf = Replace(Buf, "l?quidos", "líquidos")
   Buf = Replace(Buf, "peque?os", "pequeños")
   Buf = Replace(Buf, "verduler?as", "verdulerías")
   Buf = Replace(Buf, "botiller?as", "botillerías")
   Buf = Replace(Buf, "m?sica", "música")
   Buf = Replace(Buf, "ortop?dicos", "ortopédicos")
   Buf = Replace(Buf, "joyer?a", "joyería")
   Buf = Replace(Buf, "relojer?a", "relojería")
   Buf = Replace(Buf, "le?a", "leña")
   Buf = Replace(Buf, "artesan?as", "artesanías")
   Buf = Replace(Buf, "antig?edades", "antigüedades")
   Buf = Replace(Buf, "v?a", "vía")
   Buf = Replace(Buf, "telef?nica", "telefónica")
   Buf = Replace(Buf, "locomoci?n", "locomoción")
   Buf = Replace(Buf, "tuber?as", "tuberías")
   Buf = Replace(Buf, "mar?timo", "marítimo")
   Buf = Replace(Buf, "v?as", "vías")
   Buf = Replace(Buf, "a?rea", "aérea")
   Buf = Replace(Buf, "frigor?ficos", "frigoríficos")
   Buf = Replace(Buf, "dep?sito", "depósito")
   Buf = Replace(Buf, "parqu?metros", "parquímetros")
   Buf = Replace(Buf, "acu?tico", "acuático")
   Buf = Replace(Buf, "a?reo", "aéreo")
   Buf = Replace(Buf, "Manipulaci?n", "Manipulación")
   Buf = Replace(Buf, "mensajer?a", "mensajería")
   Buf = Replace(Buf, "m?vil", "móvil")
   Buf = Replace(Buf, "banqueter?a", "banquetería")
   Buf = Replace(Buf, "concesi?n", "concesión")
   Buf = Replace(Buf, "alimentaci?n", "alimentación")
   Buf = Replace(Buf, "Edici?n", "Edición")
   Buf = Replace(Buf, "peri?dicas", "periódicas")
   Buf = Replace(Buf, "edici?n", "edición")
   Buf = Replace(Buf, "pel?culas", "películas")
   Buf = Replace(Buf, "cinematogr?ficas", "cinematográficas")
   Buf = Replace(Buf, "televisi?n", "televisión")
   Buf = Replace(Buf, "postproducci?n", "postproducción")
   Buf = Replace(Buf, "exhibici?n", "exhibición")
   Buf = Replace(Buf, "grabaci?n", "grabación")
   Buf = Replace(Buf, "Programaci?n", "Programación")
   Buf = Replace(Buf, "Telefon?a", "Telefonía")
   Buf = Replace(Buf, "Televisi?n", "Televisión")
   Buf = Replace(Buf, "al?mbricas", "alámbricas")
   Buf = Replace(Buf, "m?viles", "móviles")
   Buf = Replace(Buf, "inal?mbrica", "inalámbrica")
   Buf = Replace(Buf, "inal?mbricas", "inalámbricas")
   Buf = Replace(Buf, "sat?lite", "satélite")
   Buf = Replace(Buf, "programaci?n", "programación")
   Buf = Replace(Buf, "inform?tica", "informática")
   Buf = Replace(Buf, "consultor?a", "consultoría")
   Buf = Replace(Buf, "inform?ticas", "informáticas")
   Buf = Replace(Buf, "tecnolog?a", "tecnología")
   Buf = Replace(Buf, "intermediaci?n", "intermediación")
   Buf = Replace(Buf, "inversi?n", "inversión")
   Buf = Replace(Buf, "cr?dito", "crédito")
   Buf = Replace(Buf, "compensaci?n", "compensación")
   Buf = Replace(Buf, "Administraci?n", "Administración")
   Buf = Replace(Buf, "asesor?a", "asesoría")
   Buf = Replace(Buf, "Evaluaci?n", "Evaluación")
   Buf = Replace(Buf, "da?os", "daños")
   Buf = Replace(Buf, "representaci?n", "representación")
   Buf = Replace(Buf, "jur?dica", "jurídica")
   Buf = Replace(Buf, "ra?ces", "raíces")
   Buf = Replace(Buf, "s?ndicos", "síndicos")
   Buf = Replace(Buf, "jur?dicas", "jurídicas")
   Buf = Replace(Buf, "tenedur?a", "teneduría")
   Buf = Replace(Buf, "auditor?a", "auditoría")
   Buf = Replace(Buf, "dise?o", "diseño")
   Buf = Replace(Buf, "t?cnica", "técnica")
   Buf = Replace(Buf, "revisi?n", "revisión")
   Buf = Replace(Buf, "an?lisis", "análisis")
   Buf = Replace(Buf, "t?cnicos", "técnicos")
   Buf = Replace(Buf, "opini?n", "opinión")
   Buf = Replace(Buf, "p?blica", "pública")
   Buf = Replace(Buf, "decoraci?n", "decoración")
   Buf = Replace(Buf, "ampliaci?n", "ampliación")
   Buf = Replace(Buf, "fotograf?as", "fotografías")
   Buf = Replace(Buf, "fotograf?a", "fotografía")
   Buf = Replace(Buf, "Asesor?a", "Asesoría")
   Buf = Replace(Buf, "peque?as", "pequeñas")
   Buf = Replace(Buf, "traducci?n", "traducción")
   Buf = Replace(Buf, "interpretaci?n", "interpretación")
   Buf = Replace(Buf, "p?blicas", "públicas")
   Buf = Replace(Buf, "cient?ficas", "científicas")
   Buf = Replace(Buf, "t?cnicas", "técnicas")
   Buf = Replace(Buf, "cl?nicas", "clínicas")
   Buf = Replace(Buf, "dotaci?n", "dotación")
   Buf = Replace(Buf, "tur?sticos", "turísticos")
   Buf = Replace(Buf, "cerrajer?a", "cerrajería")
   Buf = Replace(Buf, "investigaci?n", "investigación")
   Buf = Replace(Buf, "Desratizaci?n", "Desratización")
   Buf = Replace(Buf, "desinfecci?n", "desinfección")
   Buf = Replace(Buf, "jardiner?a", "jardinería")
   Buf = Replace(Buf, "preparaci?n", "preparación")
   Buf = Replace(Buf, "Organizaci?n", "Organización")
   Buf = Replace(Buf, "calificaci?n", "calificación")
   Buf = Replace(Buf, "administraci?n", "administración")
   Buf = Replace(Buf, "Regulaci?n", "Regulación")
   Buf = Replace(Buf, "facilitaci?n", "facilitación")
   Buf = Replace(Buf, "econ?mica", "económica")
   Buf = Replace(Buf, "Previsi?n", "Previsión")
   Buf = Replace(Buf, "afiliaci?n", "afiliación")
   Buf = Replace(Buf, "Ense?anza", "Enseñanza")
   Buf = Replace(Buf, "cient?fico", "científico")
   Buf = Replace(Buf, "t?cnico", "técnico")
   Buf = Replace(Buf, "formaci?n", "formación")
   Buf = Replace(Buf, "educaci?n", "educación")
   Buf = Replace(Buf, "ense?anza", "enseñanza")
   Buf = Replace(Buf, "atenci?n", "atención")
   Buf = Replace(Buf, "odontol?gica", "odontológica")
   Buf = Replace(Buf, "odont?logos", "odontólogos")
   Buf = Replace(Buf, "cl?nicos", "clínicos")
   Buf = Replace(Buf, "enfermer?a", "enfermería")
   Buf = Replace(Buf, "toxic?manos", "toxicómanos")
   Buf = Replace(Buf, "f?sica", "física")
   Buf = Replace(Buf, "espect?culos", "espectáculos")
   Buf = Replace(Buf, "esc?nicas", "escénicas")
   Buf = Replace(Buf, "art?sticas", "artísticas")
   Buf = Replace(Buf, "compa??as", "compañías")
   Buf = Replace(Buf, "m?sicos", "músicos")
   Buf = Replace(Buf, "hist?ricos", "históricos")
   Buf = Replace(Buf, "zool?gicos", "zoológicos")
   Buf = Replace(Buf, "Hip?dromos", "Hipódromos")
   Buf = Replace(Buf, "Gesti?n", "Gestión")
   Buf = Replace(Buf, "f?tbol", "fútbol")
   Buf = Replace(Buf, "Promoci?n", "Promoción")
   Buf = Replace(Buf, "organizaci?n", "organización")
   Buf = Replace(Buf, "tem?ticos", "temáticos")
   Buf = Replace(Buf, "pol?ticas", "políticas")
   Buf = Replace(Buf, "tel?fonos", "teléfonos")
   Buf = Replace(Buf, "Peluquer?a", "Peluquería")
   Buf = Replace(Buf, "guarder?a", "guardería")
   Buf = Replace(Buf, "peluquer?a", "peluquería")
   Buf = Replace(Buf, "ba?os", "baños")
   Buf = Replace(Buf, "p?blicos", "públicos")
   Buf = Replace(Buf, "?rganos", "órganos")
   Buf = Replace(Buf, "Seg?n", "Según")
   Buf = Replace(Buf, "Centralizaci?n", "Centralización")
   Buf = Replace(Buf, "Correcci?n", "Corrección")
   Buf = Replace(Buf, "Contabilizaci?n", "Contabilización")
   Buf = Replace(Buf, "Declaraci?n", "Declaración")
   Buf = Replace(Buf, "A?o", "Año")
   Buf = Replace(Buf, "Depreciaci?n", "Depreciación")
   Buf = Replace(Buf, "B?sicos", "Básicos")
   Buf = Replace(Buf, "a?o", "año")
   Buf = Replace(Buf, "Dep?sitos", "Depósitos")
   Buf = Replace(Buf, "Tr?nsito", "Tránsito")
   Buf = Replace(Buf, "Burs?til", "Bursátil")
   Buf = Replace(Buf, "inter?s", "interés")
   Buf = Replace(Buf, "Cr?dito", "Crédito")
   Buf = Replace(Buf, "Estimaci?n", "Estimación")
   Buf = Replace(Buf, "Gratificaci?n", "Gratificación")
   Buf = Replace(Buf, "Pr?stamos", "Préstamos")
   Buf = Replace(Buf, "Garant?a", "Garantía")
   Buf = Replace(Buf, "Mercader?as", "Mercaderías")
   Buf = Replace(Buf, "Provisi?n", "Provisión")
   Buf = Replace(Buf, "Biol?gicos", "Biológicos")
   Buf = Replace(Buf, "biol?gicos", "biológicos")
   Buf = Replace(Buf, "Cr?ditos", "Créditos")
   Buf = Replace(Buf, "m?todo", "método")
   Buf = Replace(Buf, "participaci?n", "participación")
   Buf = Replace(Buf, "L?neas", "Líneas")
   Buf = Replace(Buf, "Telef?nicas", "Telefónicas")
   Buf = Replace(Buf, "Veh?culos", "Vehículos")
   Buf = Replace(Buf, "Autom?viles", "Automóviles")
   Buf = Replace(Buf, "Biol?gico", "Biológico")
   Buf = Replace(Buf, "Inversi?n", "Inversión")
   Buf = Replace(Buf, "Porci?n", "Porción")
   Buf = Replace(Buf, "?nico", "Único")
   Buf = Replace(Buf, "Cesant?a", "Cesantía")
   Buf = Replace(Buf, "Participaci?n", "Participación")
   Buf = Replace(Buf, "D?bito", "Débito")
   Buf = Replace(Buf, "A?os", "Años")
   Buf = Replace(Buf, "disposici?n", "disposición")
   Buf = Replace(Buf, "P?rdidas", "Pérdidas")
   Buf = Replace(Buf, "P?rdida", "Pérdida")
   Buf = Replace(Buf, "Emisi?n", "Emisión")
   Buf = Replace(Buf, "conversi?n", "conversión")
   Buf = Replace(Buf, "p?rdidas", "pérdidas")
   Buf = Replace(Buf, "Prestaci?n", "Prestación")
   Buf = Replace(Buf, "Asignaci?n", "Asignación")
   Buf = Replace(Buf, "Movilizaci?n", "Movilización")
   Buf = Replace(Buf, "Colaci?n", "Colación")
   Buf = Replace(Buf, "Vi?ticos", "Viáticos")
   Buf = Replace(Buf, "Capacitaci?n", "Capacitación")
   Buf = Replace(Buf, "Mantenci?n", "Mantención")
   Buf = Replace(Buf, "Remodelaci?n", "Remodelación")
   Buf = Replace(Buf, "T?cnicos", "Técnicos")
   Buf = Replace(Buf, "Calefacci?n", "Calefacción")
   Buf = Replace(Buf, "Tel?fono", "Teléfono")
   Buf = Replace(Buf, "Cafeter?a", "Cafetería")
   Buf = Replace(Buf, "Suscripci?n", "Suscripción")
   Buf = Replace(Buf, "Art?culos", "Artículos")
   Buf = Replace(Buf, "Papeler?a", "Papelería")
   Buf = Replace(Buf, "Enajenaci?n", "Enajenación")
   Buf = Replace(Buf, "p?rdida", "pérdida")
   Buf = Replace(Buf, "Categor?a", "Categoría")
   Buf = Replace(Buf, "D?lar", "Dólar")
   Buf = Replace(Buf, "Consolidaci?n", "Consolidación")
   Buf = Replace(Buf, "Retasaci?n", "Retasación")
   Buf = Replace(Buf, "T?cnica", "Técnica")
   Buf = Replace(Buf, "Retazaci?n", "Retazación")
   Buf = Replace(Buf, "Relaci?n", "Relación")
   Buf = Replace(Buf, "Amortizaci?n", "Amortización")
   Buf = Replace(Buf, "Ejecuci?n", "Ejecución")
   Buf = Replace(Buf, "pr?stamos", "préstamos")
   Buf = Replace(Buf, "Investigaci?n", "Investigación")
   Buf = Replace(Buf, "Indemnizaci?n", "Indemnización")
   Buf = Replace(Buf, "revalorizaci?n", "revalorización")
   Buf = Replace(Buf, "Aplicaci?n", "Aplicación")
   Buf = Replace(Buf, "Realizaci?n", "Realización")
   Buf = Replace(Buf, "enajenaci?n", "enajenación")
   Buf = Replace(Buf, "Raz?n", "Razón")
   Buf = Replace(Buf, "?cida", "Ácida")
   Buf = Replace(Buf, "Tesorer?a", "Tesorería")
   Buf = Replace(Buf, "M?rgen", "Márgen")
   Buf = Replace(Buf, "L?quida", "Líquida")
   Buf = Replace(Buf, "Rotaci?n", "Rotación")
   Buf = Replace(Buf, "D?as", "Días")
   Buf = Replace(Buf, "CAMI?A", "CAMIÑA")
   Buf = Replace(Buf, "CHA?ARAL", "CHAÑARAL")
   Buf = Replace(Buf, "VICU?A", "VICUÑA")
   Buf = Replace(Buf, "VI?A", "VIÑA")
   Buf = Replace(Buf, "DO?IHUE", "DOÑIHUE")
   Buf = Replace(Buf, "HUALA?E", "HUALAÑE")
   Buf = Replace(Buf, "CA?ETE", "CAÑETE")
   Buf = Replace(Buf, "IBA?EZ", "IBAÑEZ")
   Buf = Replace(Buf, "PE?AFLOR", "PEÑAFLOR")
   Buf = Replace(Buf, "?U?OA", "ÑUÑOA")
   Buf = Replace(Buf, "PE?ALOLEN", "PEÑALOLEN")
   Buf = Replace(Buf, "Liquidaci?n", "Liquidación")
   Buf = Replace(Buf, "Devoluci?n", "Devolución")
   Buf = Replace(Buf, "Exportaci?n", "Exportación")
   Buf = Replace(Buf, "M?quina", "Máquina")
   Buf = Replace(Buf, "Electr?nico", "Electrónico")
   Buf = Replace(Buf, "Remuneraci?n", "Remuneración")
   Buf = Replace(Buf, "Champa?a", "Champaña")
   Buf = Replace(Buf, "Alcoh?licas", "Alcohólicas")
   Buf = Replace(Buf, "Analcoh?licas", "Analcohólicas")
   Buf = Replace(Buf, "D?b", "Déb")
   Buf = Replace(Buf, "Cr?d", "Créd")
   Buf = Replace(Buf, "Espec?fico", "Específico")
   Buf = Replace(Buf, "Az?car", "Azúcar")
   Buf = Replace(Buf, "Hidrobiol?gicas", "Hidrobiológicas")
   Buf = Replace(Buf, "Comercializaci?n", "Comercialización")
   Buf = Replace(Buf, "Petr?leo", "Petróleo")
   Buf = Replace(Buf, "analcoh?licas", "analcohólicas")
   Buf = Replace(Buf, "Retenci?n", "Retención")
   Buf = Replace(Buf, "m?rgen", "márgen")
   Buf = Replace(Buf, "comercializaci?n", "comercialización")
   Buf = Replace(Buf, "Analcoh?l", "Analcohól")
    CorrigeTexto = True
End Function
