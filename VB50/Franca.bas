Attribute VB_Name = "Franca"
Option Explicit

Public Const LDBLFMT3 = "#,##0.000"
Public Const LDBLFMT = "#,##0.00"
Public Const XDBLFMT = "#,##0.00####"

Public Const BL_NUMFMT = "#,###"
Public Const BL_DBLFMT1 = "#,###.#"
Public Const BL_DBLFMT2 = "#,###.##"

Public Const VLNUMFMT = "0.000000000000000"

Public Const F_SHORTDATE = "dd/mm/yy"

Dim OpenRs As Integer

Public gCalcApp As String

Public Type FontDef_t
   FontName As String
   FontSize As Single
   FontBold As Boolean
   FontUnderline As Boolean
End Type

Sub KeyTime(KeyAscii As Integer)
   Dim ch As String

   ch = Chr$(KeyAscii)
   
   If KeySys(KeyAscii) = False And Not IsNumeric(ch) And ch <> ":" Then
      Beep
      KeyAscii = 0
   End If

End Sub
Public Function ValHora(ByVal StrHora As String, DblHora As Double) As Integer
   Dim H As Double
   Dim d As Date
   
   H = 0
   On Error Resume Next
   d = Format(StrHora)
   H = d
   On Error GoTo 0
   
   If H = 0 Or H >= 24 Then
      ValHora = False
      DblHora = 0
      Exit Function
   ElseIf H >= 1 Then
      H = H * TimeSerial(1, 0, 0)
   End If
   
   DblHora = H
   ValHora = True

End Function

Public Function Num2SQL(n As Double)
   Num2SQL = "0" & Trim(str(Format(n, VLNUMFMT)))
End Function
'formatea fechas con nombre del día
Public Function LFmtDate(ByVal LDate As Long) As String
   LFmtDate = Format(LDate, "ddd dd mmm yyyy")
End Function
Public Function PrepararPrt(CmPrtDlg As Control) As Boolean
   Dim Prt As Printer

   PrepararPrt = False
   
   On Error Resume Next
   
   Printer.TrackDefault = True      ' ** PAM 10 DIC 2004
   CmPrtDlg.PrinterDefault = True   ' ** PAM 10 DIC 2004
   
   CmPrtDlg.Flags = cdlPDPrintSetup
   CmPrtDlg.CancelError = True
   CmPrtDlg.ShowPrinter
         
   If Err Then
      If Err <> cdlCancel Then
         MsgBox "Error " & Err & ", " & Error, vbExclamation
      End If
      Exit Function
   End If

   On Error GoTo 0

   Printer.EndDoc
      
   Printer.Orientation = CmPrtDlg.Orientation

   Debug.Print "Printer " & Printer.DeviceName & ": Orient=" & Printer.Orientation & ", PageSize=" & Printer.PaperSize & ", Width=" & Printer.Width

   DoEvents

   PrepararPrt = True

End Function
'Separa el string FromStr en N trozos de NumChar caracteres, asigándolos en el arreglo ToStr()
'Retorna el número de trozos encontrados
Public Function SplitStr(ByVal FromStr As String, ByVal NumChar As Integer, ToStr() As String) As Integer
   Dim i As Integer
   Dim Aux As String
   Dim j As Integer
     
   i = 0
   
   Aux = FromStr & " "
   
   If Trim(Aux) <> "" Then
      Aux = ReplaceStr(Aux, NL, " ")
      Aux = ReplaceStr(Aux, CR, "")
   End If
   
   Do While Aux <> ""
      j = Len(Aux)
      
      If j < NumChar Then
         ToStr(i) = Aux
         
         Aux = ""
      Else
         j = NumChar + 1
         Do While Mid(Aux, j, 1) <> " "
            j = j - 1
            If j <= 0 Then
               j = NumChar
               Exit Do
            End If
         Loop
         
         ToStr(i) = Left(Aux, j)
         
         Aux = Mid(Aux, j + 1)
      End If
   
      i = i + 1
   Loop
   
   SplitStr = i
   
   'ToStr(i) = Left(FromStr, NumChar)
   'RightStr = Trim(Mid(FromStr, NumChar + 1))
   
   'Do While RightStr <> ""
   '   i = i + 1
   '   ToStr(i) = Left(RightStr, NumChar)
   '   RightStr = Trim(Mid(RightStr, NumChar + 1))
   'Loop
   
   'SplitStr = i + 1
   
End Function
'Separa el string FromStr en N trozos, cortándolos en cada Return, asigándolos en el arreglo ToStr()
'Retorna el número de trozos encontrados
Public Function SplitStrReturn(ByVal FromStr As String, ToStr() As String) As Integer
   Dim i As Integer
   Dim Aux As String
   Dim j As Integer
     
   i = 0
   
   Aux = FromStr
      
   Do While Aux <> ""
      j = InStr(1, Aux, CRNL)
   
      If j > 0 Then
         ToStr(i) = Left(Aux, j - 1)
         Aux = Mid(Aux, j + 2)
      Else
         ToStr(i) = Aux
         Aux = ""
      End If
         
      i = i + 1
   Loop
   
   SplitStrReturn = i
      
End Function

'rellena con caracteres a la derecha
Public Function FillStrR(ByVal StrValue As String, ByVal StrLen As Integer) As String
   Dim Txt As String
   
   Txt = StrValue

   If Trim(Txt) <> "" Then
      Txt = ReplaceStr(Txt, Chr$(10), " ")
      Txt = ReplaceStr(Txt, Chr$(13), "")
   End If

   FillStrR = Right(String(StrLen, " ") & Txt, StrLen)

End Function

Sub KeyUserId(KeyAscii As Integer)
   Dim ch As String * 1
   
   Call KeyLower(KeyAscii)

   ch = Chr(KeyAscii)

   If KeySys(KeyAscii) = False And KeyAscii <> vbKeyReturn And Not IsNumeric(ch) And (ch < "a" Or ch > "z") And ch <> "-" And ch <> "_" Then
      Beep
      KeyAscii = 0
   End If
End Sub

Function ShortName(ByVal StrName As String) As String
   Dim ShName As String
      
   ShName = UCase(StrName)
   ShName = ReplaceStr(ShName, " DE ", " ")
   ShName = ReplaceStr(ShName, "ADMINISTRADORA", "ADM.")
   ShName = ReplaceStr(ShName, "ASOCIACION", "ASOC.")
   ShName = ReplaceStr(ShName, "BANCO", "BCO.")
   ShName = ReplaceStr(ShName, "CHILENA", "CH.")
   ShName = ReplaceStr(ShName, "CLINICA", "CL.")
   ShName = ReplaceStr(ShName, "COMERCIALIZADORA", "COM.")
   ShName = ReplaceStr(ShName, "COMERCIAL", "COM.")
   ShName = ReplaceStr(ShName, "COMPAÑIA", "CIA.")
   ShName = ReplaceStr(ShName, "CONSORCIO", "CONS.")
   ShName = ReplaceStr(ShName, "COOPERATIVA", "COOP.")
   ShName = ReplaceStr(ShName, "CORPORACION", "CORP.")
   ShName = ReplaceStr(ShName, "CORREDORA", "CORR.")
   ShName = ReplaceStr(ShName, "CORREDORES", "CORR.")
   ShName = ReplaceStr(ShName, "DIRECCION", "DIR.")
   ShName = ReplaceStr(ShName, "EMPRESAS", "EMPS.")
   ShName = ReplaceStr(ShName, "EMPRESA", "EMP.")
   ShName = ReplaceStr(ShName, "FUNDACION", "FUND.")
   ShName = ReplaceStr(ShName, "GENERALES", "GRALES.")
   ShName = ReplaceStr(ShName, "GENERAL", "GRAL.")
   ShName = ReplaceStr(ShName, "HOSPITAL", "HOSP.")
   ShName = ReplaceStr(ShName, "IMPORTADORA", "IMP.")
   ShName = ReplaceStr(ShName, "INMOBILIARIA", "INMOB.")
   ShName = ReplaceStr(ShName, "LABORATORIOS", "LABS.")
   ShName = ReplaceStr(ShName, "LABORATORIO", "LAB.")
   ShName = ReplaceStr(ShName, "LIMITADA", "LTDA.")
   ShName = ReplaceStr(ShName, "NACIONAL", "NAC.")
   ShName = ReplaceStr(ShName, "SOCIEDAD", "SOC.")
   
   ShortName = FCase(ShName)
   
End Function

'determina si un string es un patrón de búsqueda para SQL
'es decir, si tiene *, #, ?
Public Function IsPattern(Pat As String) As Boolean
   
   IsPattern = False

   If InStr(Pat, "*") > 0 Then
      IsPattern = True
   ElseIf InStr(Pat, "?") > 0 Then
      IsPattern = True
   ElseIf InStr(Pat, "#") > 0 Then
      IsPattern = True
   End If
   
End Function
Public Function LineGetStr(ByVal Buf As String, i As Integer, ByVal ErrMsg As String, ByVal NullMsg As String, StrVal As String, Optional EndLinea As Boolean = False) As Boolean
   Dim j As Integer
   
   LineGetStr = False
   StrVal = ""
   
   If Not EndLinea Then
      j = InStr(i, Buf, tb)
      
      If j = 0 Then
         MsgBeep vbExclamation
         MsgBox ErrMsg & " Tab de separación de columnas no encontrado.", vbExclamation
         Exit Function
      End If
   
      StrVal = Trim(Mid(Buf, i, j - i))
   Else
      StrVal = Trim(Mid(Buf, i))
   End If
      
   If StrVal = "" And NullMsg <> "" Then
      MsgBeep vbExclamation
      MsgBox ErrMsg & NullMsg, vbExclamation
      Exit Function
   End If
   
   i = j + 1

   LineGetStr = True
End Function
Public Function LineGetVal(ByVal Buf As String, i As Integer, ByVal ErrMsg As String, ByVal NullMsg As String, DblVal As Variant, Optional EndLinea As Boolean = False) As Boolean
   Dim j As Integer
   Dim Aux As String
   
   LineGetVal = False
   DblVal = 0
      
   If Not EndLinea Then
      j = InStr(i, Buf, tb)
      
      If j = 0 Then
         MsgBeep vbExclamation
         MsgBox ErrMsg & " Tab de separación de columnas no encontrado (" & Mid(Buf, i) & ").", vbExclamation
         Exit Function
      End If
   
      Aux = Trim(Mid(Buf, i, j - i))
   Else
      Aux = Trim(Mid(Buf, i))
   End If
   
   DblVal = Val(Aux)

   If Aux = "" And NullMsg <> "" Then
      MsgBeep vbExclamation
      MsgBox ErrMsg & NullMsg, vbExclamation
      Exit Function
   End If
   
   i = j + 1
   
   LineGetVal = True
End Function

Function PrtVertical() As Integer
   Dim PrtOrient As Integer
   
        On Error Resume Next

   PrtOrient = Printer.Orientation
   
   If Printer.Orientation <> cdlPortrait Then
      Printer.Orientation = cdlPortrait   'cdlportrait
   End If

   PrtVertical = PrtOrient

End Function

Function PrtHorizontal() As Integer
   Dim PrtOrient As Integer
 
        On Error Resume Next

   PrtOrient = Printer.Orientation
   
   If Printer.Orientation <> cdlLandscape Then
      Printer.Orientation = cdlLandscape
   End If

   PrtHorizontal = PrtOrient

End Function

Public Function GetFileExt(ByVal FName As String) As String
   Dim i As Integer
   
   i = Len(FName)
   
   Do While i > 0
      If Mid(FName, i, 1) = "." Then
         Exit Do
      End If
      
      i = i - 1
   
   Loop
      
   If i > 0 Then
      GetFileExt = Mid(FName, i)
   Else
      GetFileExt = ""
   End If
   
End Function
Public Sub Calculadora()
   Dim CalcApp As String
   Dim i As Integer
   
   If gCalcApp <> "" Then
      CalcApp = gCalcApp
   Else
      CalcApp = "calc.exe"
   End If
   
   On Error Resume Next
   
   Shell CalcApp, vbNormalFocus
   If Err Then
      MsgErr CalcApp
   End If
   
End Sub

Public Function LFill0Str(ByVal str As String, ByVal StrLen As Integer, Optional ByVal FillChar As String = "0") As String
   
   LFill0Str = Right(String(StrLen, FillChar) & str, StrLen)
      
End Function

'reconoce fechas de la forma "dd/mm/yy" o "dd-mm-yy"
Function vFmtTxtDate(ByVal StrDate As String) As Long
   Dim Mes As Integer
   Dim Dia As Integer
   Dim Ano As Integer
   
   vFmtTxtDate = 0
   If Len(StrDate) <> 8 Then
      Exit Function
   End If
   
   Dia = Val(Left(StrDate, 2))
   Mes = Val(Mid(StrDate, 4, 2))
   Ano = Val(Mid(StrDate, 7, 2))
   
   If Dia = 0 Or Mes = 0 Or Ano = 0 Then
      Exit Function
   End If
   
   vFmtTxtDate = DateSerial(Ano, Mes, Dia)
   
End Function
Public Sub AppendWhere(ByVal NewCond As String, Where As String, Optional ByVal WOper As String = "AND")
   
   If Trim(NewCond) = "" Then
      Exit Sub
   End If
   
   If Where = "" Then
      Where = " WHERE " & NewCond
   Else
      Where = Where & " " & WOper & " " & NewCond
   End If

End Sub

Public Function GetPrtTextWidth(ByVal Txt As String, Optional ByVal bMsg As Boolean = 1) As Integer
   Static bNoImp As Boolean
   Dim Msg As String

   On Error Resume Next

   If Printer Is Nothing Then
      
      If bNoImp = False And bMsg Then
         MsgBox1 "No se detectó ninguna impresora.", vbExclamation
         bNoImp = True
      End If
      
      Exit Function
   End If

   GetPrtTextWidth = Printer.TextWidth(Txt)
   
   If Err.Number Then
      Msg = "Problemas en la comunicación con la impresora """ & Printer.DeviceName & """." & vbCrLf & "Error " & Err.Number & ", " & Err.Description
      Call AddLog("GetPrtTextWidth: " & Replace(Msg, vbCrLf, " "))
      If bMsg Then
         MsgErr "ATENCIÓN" & vbCrLf & Msg & vbCrLf & "Pruebe a cambiar la impresora por defecto de Windows.", vbExclamation
      End If
      
      GetPrtTextWidth = Len(Txt) * (1350 / 17) ' estimación
   End If

End Function
Public Function StringStr(ByVal Number As Long, FillStr As String) As String
   Dim str As String
   Dim i As Integer
   
   For i = 1 To Number
      str = str & FillStr
   Next i
      
   StringStr = str
End Function

Public Function FillCbStrArray(Cb As ComboBox, StrArr() As String, Optional ByVal CbClear As Boolean = True, Optional ByVal AddBlank As Boolean = False)
   Dim i As Integer
   
   If CbClear Then
      Cb.Clear
   End If
   
   If AddBlank Then
      Call CbAddItem(Cb, " ", 0)
   End If
   
   For i = 1 To UBound(StrArr)
                If Len(StrArr(i)) > 0 Then ' 19 oct 2016: pam
                        Call CbAddItem(Cb, StrArr(i), i)
                End If
   Next i
   
End Function
Public Function ReplaceStrStartAt(ByVal Where As String, ByVal FromStr As String, ByVal ToStr As String, ByVal StartAt As Integer, Optional ByVal compare As VbCompareMethod = vbTextCompare) As String
   Dim Aux As String
   
   If StartAt <= 0 Then
      ReplaceStrStartAt = ""
   End If
   
   Aux = Mid(Where, StartAt)
   
   Aux = ReplaceStr(Aux, FromStr, ToStr, compare)
   
   If StartAt > 1 Then
      ReplaceStrStartAt = Left(Where, StartAt - 1) & Aux
   Else
      ReplaceStrStartAt = Aux
   End If
   

End Function

'Función para convertir de US-ASCII a Latin1 o, lo que es lo mismo iso-8859-1
'Tabla de conversión: http://string-functions.com/encodingtable.aspx?encoding=20127&decoding=28591

Public Function AsciiToLatin1(ByVal Buf As String) As String
   Dim c As String * 1
   Dim Aux As String
   Dim i As Integer
   Dim NAscii As Integer
   Dim Base As Integer
   Static Latin1CharSet(255, 1) As String
   
   If Buf = "" Then
      AsciiToLatin1 = ""
      Exit Function
   End If
      
      
   If Latin1CharSet(0, 0) = "" Then
   
      Base = 161   'Antes de este Ascii están los caracteres normales y los que no se representan en Ascii
       
      For i = 0 To Base - 1  'caracteres normales y los que no se representan en Ascii, los dejamos igual
         Latin1CharSet(i, 0) = Chr(i):  Latin1CharSet(i, 1) = Chr(i)
      Next i
      
      For i = Base To 191    'caracteres raros de Ascii los dejamos en blanco
         Latin1CharSet(i, 0) = Chr(i):  Latin1CharSet(i, 1) = " "
      Next i
      
      i = 192
      Latin1CharSet(i, 0) = "À":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Á":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Â":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Ã":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Ä":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Å":  Latin1CharSet(i, 1) = "A"
      i = i + 1
      Latin1CharSet(i, 0) = "Æ":  Latin1CharSet(i, 1) = "A"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ç":  Latin1CharSet(i, 1) = "C"
      i = i + 1
      Latin1CharSet(i, 0) = "È":  Latin1CharSet(i, 1) = "E"
      i = i + 1
      Latin1CharSet(i, 0) = "É":  Latin1CharSet(i, 1) = "E"
      i = i + 1
      Latin1CharSet(i, 0) = "Ê":  Latin1CharSet(i, 1) = "E"
      i = i + 1
      Latin1CharSet(i, 0) = "Ë":  Latin1CharSet(i, 1) = "E"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ì":  Latin1CharSet(i, 1) = "I"
      i = i + 1
      Latin1CharSet(i, 0) = "Í":  Latin1CharSet(i, 1) = "I"
      i = i + 1
      Latin1CharSet(i, 0) = "Î":  Latin1CharSet(i, 1) = "I"
      i = i + 1
      Latin1CharSet(i, 0) = "Ï":  Latin1CharSet(i, 1) = "I"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ð":  Latin1CharSet(i, 1) = "D"
      i = i + 1
      Latin1CharSet(i, 0) = "Ñ":  Latin1CharSet(i, 1) = "N"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ò":  Latin1CharSet(i, 1) = "O"
      i = i + 1
      Latin1CharSet(i, 0) = "Ó":  Latin1CharSet(i, 1) = "O"
      i = i + 1
      Latin1CharSet(i, 0) = "Ô":  Latin1CharSet(i, 1) = "O"
      i = i + 1
      Latin1CharSet(i, 0) = "Õ":  Latin1CharSet(i, 1) = "O"
      i = i + 1
      Latin1CharSet(i, 0) = "Ö":  Latin1CharSet(i, 1) = "O"
      
      i = i + 1
      Latin1CharSet(i, 0) = "×":  Latin1CharSet(i, 1) = " "
      i = i + 1
      Latin1CharSet(i, 0) = "Ø":  Latin1CharSet(i, 1) = "0"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ù":  Latin1CharSet(i, 1) = "U"
      i = i + 1
      Latin1CharSet(i, 0) = "Ú":  Latin1CharSet(i, 1) = "U"
      i = i + 1
      Latin1CharSet(i, 0) = "Û":  Latin1CharSet(i, 1) = "U"
      i = i + 1
      Latin1CharSet(i, 0) = "Ü":  Latin1CharSet(i, 1) = "U"
      
      i = i + 1
      Latin1CharSet(i, 0) = "Ý":  Latin1CharSet(i, 1) = "Y"
      i = i + 1
      Latin1CharSet(i, 0) = "Þ":  Latin1CharSet(i, 1) = "P"
      i = i + 1
      Latin1CharSet(i, 0) = "ß":  Latin1CharSet(i, 1) = "B"
      
      i = i + 1
      Latin1CharSet(i, 0) = "à":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "á":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "â":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "ã":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "ä":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "å":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "æ":  Latin1CharSet(i, 1) = "a"
      i = i + 1
      Latin1CharSet(i, 0) = "ç":  Latin1CharSet(i, 1) = "c"
      
      i = i + 1
      Latin1CharSet(i, 0) = "è":  Latin1CharSet(i, 1) = "e"
      i = i + 1
      Latin1CharSet(i, 0) = "é":  Latin1CharSet(i, 1) = "e"
      i = i + 1
      Latin1CharSet(i, 0) = "ê":  Latin1CharSet(i, 1) = "e"
      i = i + 1
      Latin1CharSet(i, 0) = "ë":  Latin1CharSet(i, 1) = "e"
      
      i = i + 1
      Latin1CharSet(i, 0) = "ì":  Latin1CharSet(i, 1) = "i"
      i = i + 1
      Latin1CharSet(i, 0) = "í":  Latin1CharSet(i, 1) = "i"
      i = i + 1
      Latin1CharSet(i, 0) = "î":  Latin1CharSet(i, 1) = "i"
      i = i + 1
      Latin1CharSet(i, 0) = "ï":  Latin1CharSet(i, 1) = "i"
      
      i = i + 1
      Latin1CharSet(i, 0) = "ð":  Latin1CharSet(i, 1) = " "
      i = i + 1
      Latin1CharSet(i, 0) = "ñ":  Latin1CharSet(i, 1) = "n"
        
      
      i = i + 1
      Latin1CharSet(i, 0) = "ò":  Latin1CharSet(i, 1) = "o"
      i = i + 1
      Latin1CharSet(i, 0) = "ó":  Latin1CharSet(i, 1) = "o"
      i = i + 1
      Latin1CharSet(i, 0) = "ô":  Latin1CharSet(i, 1) = "o"
      i = i + 1
      Latin1CharSet(i, 0) = "õ":  Latin1CharSet(i, 1) = "o"
      i = i + 1
      Latin1CharSet(i, 0) = "ö":  Latin1CharSet(i, 1) = "o"
      
      i = i + 1
      Latin1CharSet(i, 0) = "÷":  Latin1CharSet(i, 1) = " "
      i = i + 1
      Latin1CharSet(i, 0) = "ø":  Latin1CharSet(i, 1) = " "
      
      i = i + 1
      Latin1CharSet(i, 0) = "ù":  Latin1CharSet(i, 1) = "u"
      i = i + 1
      Latin1CharSet(i, 0) = "ú":  Latin1CharSet(i, 1) = "u"
      i = i + 1
      Latin1CharSet(i, 0) = "û":  Latin1CharSet(i, 1) = "u"
      i = i + 1
      Latin1CharSet(i, 0) = "ü":  Latin1CharSet(i, 1) = "u"
      
      i = i + 1
      Latin1CharSet(i, 0) = "ý":  Latin1CharSet(i, 1) = "y"
      i = i + 1
      Latin1CharSet(i, 0) = "þ":  Latin1CharSet(i, 1) = "b"
      i = i + 1
      Latin1CharSet(i, 0) = "ÿ":  Latin1CharSet(i, 1) = "y"
   End If
   
   For i = 1 To Len(Buf)
      c = Mid(Buf, i, 1)
      NAscii = Asc(c)
      If NAscii < 31 Then ' 19 feb 2021 - pam: si son de control se cambia por blanco
         Aux = Aux & " "
      ElseIf NAscii >= Base Then
         Aux = Aux & Latin1CharSet(NAscii, 1)
      Else
         Aux = Aux & c
      End If
   Next i
   
   AsciiToLatin1 = Aux
End Function

Public Function ConvUnix2DosFile(ByVal FNameIn As String, ByVal FNameOut As String) As Integer
   Dim objFSO As FileSystemObject
   Dim objText As TextStream
   Dim strText As String
   Dim strNewText As String

   On Error Resume Next
   Set objFSO = New FileSystemObject
   Set objText = objFSO.OpenTextFile(FNameIn, ForReading, False, TristateUseDefault)
   
   If Err Then
      MsgErr FNameIn
      ConvUnix2DosFile = -Err
      Exit Function
   End If

   strText = objText.ReadAll()
   
   If Err Then
      MsgErr FNameIn
      ConvUnix2DosFile = -Err
      Exit Function
   End If
   
   Call objText.Close
   
   'primero vemos si hay CrLf
   If InStr(strText, vbCrLf) > 0 Then   'ya está en formato DOS
      ConvUnix2DosFile = 0
      Exit Function
   End If

   strNewText = Replace(strText, vbLf, vbCrLf)

   Set objText = objFSO.CreateTextFile(FNameOut, True)
   
   If Err Then
      MsgErr FNameOut
      ConvUnix2DosFile = -Err
      Exit Function
   End If

   Call objText.Write(strNewText)
   
   If Err Then
      MsgErr FNameOut
      ConvUnix2DosFile = -Err
      Exit Function
   End If

   Call objText.Close
   
   ConvUnix2DosFile = vbOK
   
   On Error GoTo 0

End Function
