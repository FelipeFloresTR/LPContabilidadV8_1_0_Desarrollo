Attribute VB_Name = "ModDownLast"
Option Explicit


Public Sub FwDownLast(ByVal Frm As Form, CmDialog As CommonDialog, ByVal bSiempreDemo As Boolean)
   Dim FName As String, Url As String, Params As String, FSize As Long, Buf As String, Page As String
   Dim cMd5 As ClsMd5, Md5 As String, Md5b As String, FPath As String, LastVer As String, i As Integer
   Dim Rand As Long

'   2 feb 2016: para que pueda bajar el actualizador desde la Demo si está al día
'   If gAppCode.Demo Then
'      If MsgBox1("Estimado Cliente" & vbCrLf & vbCrLf & "Para poder bajar la actualización usted debe estar con la mantención al día." & vbCrLf & "Si no es el caso, por favor comuníquese con su proveedor para averiguar cómo ponerse al día." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbYesNo Or vbInformation) <> vbYes Then
'         Exit Sub
'      End If
'   End If
   
   If Trim(gAppCode.Rut) = "" Then
      MsgBox1 "Antes de realizar la descarga, debe ingresar el RUT del dueño de la licencia en la opción Datos Oficina.", vbExclamation
      Exit Sub
   End If
   
'   Debug.Print FwEncrypt1("    http://www.fairware.cl/DownLast.asp    ", 71935)
'   Url = Trim(FwDecrypt1("92477D34ECD5BC36E4CDB8B6CABDA89C3C726C70AAAEA6B0A5A0A5B33D807B77743A81BEC6CBD164409D7B", 71935))
'   Debug.Print FwEncrypt1("    https://servicioslp.thomsonreuters.cl/DownLast.asp    ", 71935)
   Url = Trim(FwDecrypt1("90365D85AE88E04BEAC4A08F9478D4B94A71DC517CE3CCC3AC958F7D70E3D3CAC0BE31EBE4E2DAD7DBE6ECE665F272377DB9BDC5CAD0633F9C7A", 71935))
       
   Params = "Hd_ProdNom=" & APP_FULLNAME & "&Hd_Prod=u" & APP_NAME & "&Tx_RUT=" & gAppCode.Rut & "&Tx_CodPC=" & gAppCode.PcCode
   
   If bSiempreDemo Then ' 4 feb 2016
      Params = Params & "&Demo=543"
   End If
   
   Params = Params & "&k=" & GenClave("#?" & Params, 7531)
   
   Params = Params & "&r=" & Int(Rnd(Now) * 999999)
   
   Frm.MousePointer = vbHourglass
   DoEvents
   DoEvents
      
   Page = FwPostPageOld(Url, Params & "&Hd_qName=1") ' Hd_qName=1 solo consulta no descarga
   DoEvents

   If Page = "" Then
      MsgBox1 "El archivo aún no ha sido publicado, por favor intente más tarde.", vbInformation
      Frm.MousePointer = vbDefault
      Exit Sub
   End If
      
   Buf = FwGetXmlTag(Page, "errmsg")
   
   If Buf <> "" Then
      Buf = Utf8Ansi(Buf)  ' 14 abr 2019: por si venía en UTF8
      MsgBox1 ReplaceStr(Buf, "\n", vbCrLf), vbExclamation
      Frm.MousePointer = vbDefault
      Exit Sub
   End If
      
   DoEvents
      
   FName = FwGetXmlTag(Page, "lastver")
   If FName = "" Then
      MsgBox1 "No se puede descargar la actualización, inténtelo desde el sitio web.", vbExclamation
      Frm.MousePointer = vbDefault
      Exit Sub
   End If
      
   DoEvents
   
   i = Len(FName)
   LastVer = Mid(FName, i - 3 - 6, 6)   ' se asume que viene  nombreYYMMDD.exe
   
   If LastVer < Format(W.FVersion, "yymmdd") Then
      If MsgBox1("La actualización es de una versión anterior a la actual." & vbCrLf & "¿ Desea continuar ?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
         Frm.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   
   Md5 = FwGetXmlTag(Page, "md5")
   
   CmDialog.Filename = FName
   CmDialog.Filter = "Instalador (*.exe)|*.exe"
   CmDialog.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
   CmDialog.CancelError = True
   
   On Error Resume Next
   
   CmDialog.InitDir = W.DownDir
   
   CmDialog.ShowSave
   
   DoEvents

   If Err And CmDialog.Filename <> "" Then
      Frm.MousePointer = vbDefault
      Exit Sub
   End If
   
   Frm.MousePointer = vbHourglass
   DoEvents
   
'   If W.InDesign Then
'      Url = "http://localhost/fairware/DownLast.asp"
'   End If
   
   FPath = CmDialog.Filename
   FSize = FwWebSaveFile(Url, Params, FPath)
   
   If FSize > 0 Then
   
      If Md5 <> "" Then  ' si viene el md5, lo verificamos
   
         Set cMd5 = New ClsMd5
      
         Md5b = cMd5.DigestFileToHexStr(FPath)
         Set cMd5 = Nothing
      End If
      
      If Md5 = "" Or StrComp(Md5, Md5b, vbTextCompare) = 0 Then
         MsgBox1 "Se descargó el actualizador en '" & FPath & "'." & vbCrLf & "Antes de ejecutarlo, debe cerrar este programa y asegurarse que nadie más lo tenga abierto.", vbInformation
      Else
         Kill FPath
         MsgBox1 "El archivo no pudo ser descargado correctamente, intente a descargarlo desde el sitio web usando el código " & gAppCode.PcCode & ".", vbExclamation
      End If
   ElseIf FSize = -404 Then
      MsgBox1 "Archivo no encontrado " & FName & ", intente más tarde.", vbExclamation
   ElseIf FSize = -20 Then
      MsgBox1 "Pagina no encontrada, intente más tarde.", vbExclamation
   Else
      MsgBox1 "El archivo no pudo ser descargado (Err=" & gWLastDllError & ").", vbExclamation
   End If

   Frm.MousePointer = vbDefault

End Sub


