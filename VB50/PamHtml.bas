Attribute VB_Name = "PamHtml"
Option Explicit


' Para hacer los Send
Public Function HtmlEscape(ByVal Buf As String) As String
   Dim Letras As String, i As Integer
   
   Letras = "αινσϊόρ°"
   For i = 1 To Len(Letras)
      Buf = ReplaceStr(Buf, Mid(Letras, i, 1), "%" & Hex(Asc(Mid(Letras, i, 1))), vbBinaryCompare)
      Buf = ReplaceStr(Buf, UCase(Mid(Letras, i, 1)), "%" & Hex(Asc(UCase(Mid(Letras, i, 1)))), vbBinaryCompare)
   Next i

   HtmlEscape = Buf
   
End Function
Public Function HtmlEscape3(ByVal Buf As String) As String
   Dim Letras As String, i As Integer
   
   Letras = "% +=/&αινσϊόρ°"
   For i = 1 To Len(Letras)
      Buf = ReplaceStr(Buf, Mid(Letras, i, 1), "%" & Hex(Asc(Mid(Letras, i, 1))), vbBinaryCompare)
      Buf = ReplaceStr(Buf, UCase(Mid(Letras, i, 1)), "%" & Hex(Asc(UCase(Mid(Letras, i, 1)))), vbBinaryCompare)
   Next i

   HtmlEscape3 = Buf
   
End Function
' Para hacer los Send
Public Function HtmlEscape2(ByVal Buf As String, Optional ByVal bTags As Boolean = 1) As String
   Const iLT As Integer = 9
   Dim i As Integer, c As Integer, ch As String, Rep As String
   Static Letras(iLT + 1, 1) As String
   
   If Buf = "" Then
      Exit Function
   End If
      
      
   If Letras(0, 0) = "" Then
      Letras(0, 0) = "α":     Letras(0, 1) = "&aacute;"
      Letras(1, 0) = "ι":     Letras(1, 1) = "&eacute;"
      Letras(2, 0) = "ν":     Letras(2, 1) = "&iacute;"
      Letras(3, 0) = "σ":     Letras(3, 1) = "&oacute;"
      Letras(4, 0) = "ϊ":     Letras(4, 1) = "&uacute;"
      Letras(5, 0) = "ό":     Letras(5, 1) = "&uuml;"
      Letras(6, 0) = "ρ":     Letras(6, 1) = "&ntilde;"
      
      Letras(7, 0) = """":    Letras(7, 1) = "&quot;"
      Letras(8, 0) = "'":     Letras(8, 1) = "&apos;"
           
      Letras(9, 0) = "<":     Letras(9, 1) = "&lt;"
      Letras(10, 0) = ">":    Letras(10, 1) = "&gt;"
   End If

   Buf = ReplaceStr(Buf, "&", "&amp;", vbBinaryCompare) ' Primero que nada

   For i = 0 To UBound(Letras, 1)
   
'      If bTags = False And i > 6 Then  ' no toma < o >
      If bTags = False And i >= iLT Then  ' no toma < o >
         Exit For
      End If
         
      For c = 1 To 2
         ch = Letras(i, 0)
         Rep = Letras(i, 1)
         If c = 2 Then  ' Ucase
            ch = UCase(ch)
            Mid(Rep, 2, 1) = UCase(Mid(Rep, 2, 1))
         End If
   
         Buf = ReplaceStr(Buf, ch, Rep, vbBinaryCompare)
      Next c
   Next i

   HtmlEscape2 = Buf
   
End Function

