Attribute VB_Name = "PrtGrid"
Option Explicit

'definición de constantes de Align para FlexGrid que no están definidas en VB
Global Const FlxAlignLeft = flexAlignLeftCenter
Global Const FlxAlignRight = flexAlignRightCenter
Global Const FlxAlignCenter = flexAlignCenterCenter

Global Const FNT_TIMES = "Times New Roman"
Global Const FNT_MS_SANSSERIF = "MS Sans Serif"
Global Const FNT_ARIAL = "Arial"
Global Const FNT_COURIER = "Courier New"

Global Const FNT_DEFAULT = FNT_ARIAL

Global Const FOOTERSIZE = 2000   '2500

Global Const PRT_LINE_HE = 315

Global Const FOOTERPOS = FOOTERSIZE - 500

Global gPrt_GrCellHeight As Integer

Global Const PRT_GRCELLHEIGHT = 225

Sub PrtAlign_(ByVal Txt As String, ByVal X As Integer, ByVal W As Integer, ByVal Align As Integer)
   Dim Tw As Integer
   Dim Aux As String
   

   Aux = Txt
   Tw = Printer.TextWidth(Aux)
   
   Do While Tw > W
      Aux = Left(Aux, Len(Aux) - 1)
      Tw = Printer.TextWidth(Aux)
   Loop

   Select Case Align
      Case vbLeftJustify:
         Printer.CurrentX = X
      Case vbRightJustify:
         Printer.CurrentX = X + (W - Tw)
      Case vbCenter:
         Printer.CurrentX = X + (W - Tw) / 2
   End Select

   Printer.Print Aux;

End Sub
'Imprime la grilla Grid con la siguientes características:
'  Esquina superior izquierda: nombre de la empresa (Nombre)
'  Esquina superior derecha: fecha y numeración paginas
'  Centrado primera línea: Título1
'  Centrado segunda línea: Título2
'  Alineado con Título2: Título3
'  Alineado con Título2: Título4
'  A la izquierda, sobre la grilla: Encabezado
'
'Usa el font de la grilla en pantalla, a menos que UseCourier sea True.
'
'Si el alto de la fila es = 0 o el valor asociado a TextMatrix(Fila,ColObligatoria) = "" entonces no imprime la fila.
'
'Ancho de columnas igual a la pantalla, a menos que ColWi(ColObligatoria) sea <> de 0, en cuyo caso usa ColWi(i)
'
'Totales de cada columna se obtienen del arreglo Total()
'
'NTotLines indica la cantidad de lineas de totales que se desea imprimir. Estas líneas de totales se
'obtienen del arreglo Total(), suponiendo que las filas de totales se alinean a lo largo del arreglo,
'utilizando para cada línea la cantidad de columnas de la grilla. El título de cada línea de total
'se obtiene de la primera columna con ancho distinto de cero.
'
'Opcionalmente, puede haber una columna en la grilla con título ".FMT" que contenga:
'  "C(color)" en las filas que se desea imprimir con el color (color). Ejemplo: "C124" imprime la línea con el color 124
'             asegurarse que esta sea la última opción de formato de la línea (ej.: "BC124" imprime la línea en el color seleccionado y con bold
'  "B" en las filas que se desea destacar con bold
'  "L" en las filas en las que se desea insertar una línea divisoria antes de imprimir el registro
'  "LB" en las filas en las que se desea insertar una línea divisoria antes de imprimir el registro en bold
'  "T" idem a "L" salvo que la línea es entrecortada (pequeños trazos)
'
'CallEndDoc
'  Si es False (0) no llama a Printer.EndDoc
'  Si es True(-1), invoca Printer.EndDoc e imprime al final de la última página "Total pags. N"
'  Si es -2, invoca Printer.EndDoc e imprime al final de la última página "Continua >>>"
'
'PrFontName: FontName con el que se desea imprimir la grilla. Si no se pone, usa FontName de la grilla.
'Idem para FontSize.
'Obs: observaciones al final del documento
'InitPag: número primera página, cuando se imprime un doc por partes
'TitleBold: indica si Título1 es en bold. Por omisión es falso
'Retorna el número de páginas impresas
'
'+-----------------------------------------------------------+
'| Nombre                                          Fecha     |
'|                         Titulo 1                Pag.      |
'|                         Titulo 2                          |
'|                         Titulo 3                          |
'|                         Titulo 4                          |
'| Encabezado                                                |
'| +------------- grilla ----------------------------------+ |
'| |                                                       | |
'| +------------- grilla ----------------------------------+ |
'| Obs                                                          |
'+-----------------------------------------------------------+
Function PrtFlexGrid(Grid As Object, ByVal Nombre As String, ByVal Titulo1 As String, ByVal Titulo2 As String, Encabezado As String, ColWi() As Integer, Total() As String, ByVal UseCourier As Integer, Optional ColObligatoria As Integer = 1, Optional Titulo3 As String = "", Optional Titulo4 As String = "", Optional CallEndDoc As Integer = -1, Optional NTotLines As Integer = 1, Optional PrtHeader As Boolean = True, Optional PrFontName As String = "", Optional PrFontSize As Single = -1, Optional Obs As String = "", Optional InitPag As Integer = -1, Optional TitleBold As Boolean = False) As Integer
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   Dim Linea As String
   Dim Pag As Integer
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim GrRow As Integer
   Dim TLeft As Integer
   ReDim LinVert(20, 2) As Long
   Dim LeftX As Long, TopY As Long
   Dim RightX As Integer
   Dim X As Long
   Dim AuxX As Integer
   Dim fY As Single
   Dim grW As Integer
   Dim Delta As Integer
   Dim CellHeight As Integer
   Dim OfiName As String
   Dim OfiLen As Integer
   Dim FmtCol As Integer
   Dim AuxFmt As String
   Dim CurDrawStyle As Integer
   Dim AuxObs As String
   Dim AuxTopY(100) As Integer
   Dim TitWMax As Long
   Dim CurY As Long
   Dim FixedRows As Integer
   Dim FixedAlignment As Integer
   
   PrtFlexGrid = 0
   
   If Grid.Rows <= 1 Then
      MsgBeep vbExclamation
      MsgBox "Listado vacío.", vbExclamation + vbOKOnly
      Exit Function
   End If
      
   Grid.Row = 1
   grW = 0
   
   If ColWi(ColObligatoria) = 0 Then
      For i = 0 To Grid.Cols - 1
         ColWi(i) = Grid.ColWidth(i)
         If UseCourier = True Then
            ColWi(i) = ColWi(i) * 1.3
         End If
         grW = grW + ColWi(i)
      Next i
   Else
      For i = 0 To Grid.Cols - 1
         If UseCourier = True Then
            ColWi(i) = ColWi(i) * 1.3
         End If
         grW = grW + ColWi(i)
      Next i
   End If
   
   LeftX = (Printer.Width - 300 - grW) / 2  ' -300 porque a veces no imprime la línea del lado derecho
   
   fY = 0.5
   Delta = 30
   
   OldFName = Printer.FontName
   OldFBold = Printer.FontBold
   OldFSize = Printer.FontSize
   
   RightX = Printer.Width - 2000
      
   TLeft = 12
   
   If UseCourier = False Then
      'veamos si hay problema con el font
      On Error Resume Next
      Printer.FontName = FNT_DEFAULT
      Printer.FontSize = 10
      Printer.FontBold = False
      
      If Err Then
         MsgBox "Error " & Err & ", " & Error, vbExclamation
      End If
   End If
   
   FmtCol = -1
   For i = 0 To Grid.Cols - 1
      If Grid.TextMatrix(0, i) = ".FMT" Then
         FmtCol = i
      End If
   Next i
   
   If InitPag > 0 Then
      Pag = InitPag
   Else
      Pag = 1
   End If
   GrRow = Grid.FixedRows
   
   Grid.Row = 0
   Grid.Col = 0
   CellHeight = Grid.CellHeight
   If CellHeight = 0 Then
      If gPrt_GrCellHeight = 0 Then
         CellHeight = PRT_GRCELLHEIGHT
      Else
         CellHeight = gPrt_GrCellHeight
      End If
   End If
   
   Do While GrRow < Grid.Rows
            
      If PrtHeader = True Then
         'imprimimos encabezado de página
           
         Printer.CurrentY = 0
         
         Printer.Print
         Printer.Print
         
         If UseCourier = False Then
            Printer.FontName = FNT_DEFAULT
         Else
            Printer.FontName = "Courier"
         End If
         
         Printer.FontSize = 10
         Printer.Print Tab(TLeft);
         If LeftX < 0 Then
            LeftX = Printer.CurrentX
         End If
   
         Printer.FontUnderline = True
         Printer.FontSize = 18
         Printer.Print Nombre;
         Printer.FontUnderline = False
         Printer.FontSize = 8
         Printer.CurrentX = RightX
         Printer.Print " Pág. " & Pag
         
         'Printer.Print
   
         Printer.FontSize = 8
         Printer.CurrentX = RightX
         Printer.Print Format(Now, EDATEFMT)
      
         Printer.FontSize = 10
         TitWMax = Printer.TextWidth(Titulo2)
         If TitWMax < Printer.TextWidth(Titulo3) Then
            TitWMax = Printer.TextWidth(Titulo3)
         End If
         If TitWMax < Printer.TextWidth(Titulo4) Then
            TitWMax = Printer.TextWidth(Titulo4)
         End If
      
         If Titulo1 <> "" Then
            Printer.Print
            If TitleBold = True Then
               Printer.FontSize = 14
               Printer.FontBold = True
               Printer.Print Tab(TLeft);
               Printer.CurrentX = (Printer.Width - Printer.TextWidth(Titulo1)) / 2
               Printer.Print Titulo1
               Printer.FontBold = False
            Else
               Printer.FontSize = 12
               Printer.Print Tab(TLeft);
               Printer.CurrentX = (Printer.Width - Printer.TextWidth(Titulo1)) / 2
               Printer.Print Titulo1
            End If
         End If
                  
         If Titulo2 <> "" Then
            Printer.FontSize = 10
            Printer.Print Tab(TLeft);
            Printer.CurrentX = (Printer.Width - TitWMax) / 2
            Printer.Print Titulo2
         End If
         
         If Titulo3 <> "" Then
            Printer.Print Tab(TLeft);
            Printer.FontSize = 10
            Printer.CurrentX = (Printer.Width - TitWMax) / 2
            Printer.Print Titulo3
         End If
         
         If Titulo4 <> "" Then
            Printer.Print Tab(TLeft);
            Printer.FontSize = 10
            Printer.CurrentX = (Printer.Width - TitWMax) / 2
            Printer.Print Titulo4
         End If
         
         Printer.Print
         Printer.Print
         
         If Encabezado <> "" Then
            
            Printer.FontSize = 10
            i = 1
            Do
               j = InStr(i, Encabezado, CRNL)
               Printer.CurrentX = LeftX
               
               If j = 0 Then
                  Printer.Print Mid(Encabezado, i)
                  Exit Do
               Else
                  Printer.Print Mid(Encabezado, i, j - i)
                  i = j + 2
               End If
            Loop
            Printer.Print
         End If
      
      ElseIf Titulo1 <> "" Then
         Printer.Print
         Printer.Print
         
         If UseCourier = False Then
            Printer.FontName = FNT_DEFAULT
         Else
            Printer.FontName = "Courier"
         End If
         
         Printer.FontSize = 10
         Printer.Print Tab(TLeft);
         Printer.CurrentX = (Printer.Width - Printer.TextWidth(Titulo1)) / 2
         Printer.Print Titulo1
      
         Printer.Print
         Printer.Print
      End If
      
      'ponemos un font por default para la grilla
      Printer.FontSize = FNT_DEFAULT
      Printer.FontSize = 8.25
      
      If UseCourier = False Then
         If PrFontName = "" Then
            Printer.FontName = Grid.FontName
         Else
            Printer.FontName = PrFontName
         End If
         
         Printer.FontBold = False
         
         If PrFontSize = -1 Then
            Printer.FontSize = Grid.FontSize
         Else
            Printer.FontSize = PrFontSize
         End If
      Else
         Printer.FontName = "Courier"
         Printer.FontBold = False
         Printer.FontSize = 10
      End If
   
      TopY = Printer.CurrentY
      
      'imprimimos línea horizontal superior
      Printer.Line (LeftX, TopY)-(LeftX + grW + Delta, TopY)

      Printer.CurrentX = LeftX
      Printer.CurrentY = TopY + CellHeight * fY

      'imprimimos nombres de columnas
      If TypeName(Grid) = "MSFlexGrid" Then
         FixedRows = Grid.FixedRows
      Else
         FixedRows = Grid.FlxGrid.FixedRows
      End If
         
      For j = 0 To FixedRows - 1
         X = LeftX
         For i = 0 To Grid.Cols - 1
         
            If j > 0 Then
               Call PrtAlign_(Trim(Grid.TextMatrix(j, i)), X, ColWi(i), vbCenter)
               
            Else
               If TypeName(Grid) = "MSFlexGrid" Then
                  FixedAlignment = Grid.FixedAlignment(i)
               Else
                  FixedAlignment = Grid.FlxGrid.FixedAlignment(i)
               End If
         
               If FixedAlignment = flexAlignRightCenter Then
                  Call PrtAlign_(Grid.TextMatrix(j, i), X, ColWi(i), vbRightJustify)
                  AuxTopY(i) = 1
               ElseIf FixedAlignment = flexAlignLeftCenter Then
                  Call PrtAlign_(Grid.TextMatrix(j, i), X, ColWi(i), vbLeftJustify)
               Else
                  Call PrtAlign_(Grid.TextMatrix(j, i), X, ColWi(i), vbCenter)
               End If
            
            End If
            
            X = X + ColWi(i)
         Next i
         Printer.CurrentY = Printer.CurrentY + CellHeight * fY * 2
                  
      Next j
            
      'Printer.CurrentY = Printer.CurrentY + CellHeight * fY
      
      'imprimimos línea horizontal bajo nombres campos
      Printer.Line (LeftX, Printer.CurrentY + Delta)-(LeftX + grW + Delta, Printer.CurrentY + Delta)

      'imprimimos registros
      
      Do While GrRow < Grid.Rows
      
         If Grid.TextMatrix(GrRow, ColObligatoria) <> "" And Grid.RowHeight(GrRow) <> 0 Then
            
            If FmtCol >= 0 Then
               
               AuxFmt = Grid.TextMatrix(GrRow, FmtCol)
               k = 1
               Do While Mid(AuxFmt, k, 1) <> ""
                  Select Case Mid(AuxFmt, k, 1)
                     Case "L"
                        Printer.Line (LeftX, Printer.CurrentY + Delta * 3)-(LeftX + grW + Delta, Printer.CurrentY + Delta * 3)
                     Case "T"
                        CurDrawStyle = Printer.DrawStyle
                        Printer.DrawStyle = vbDash
                        Printer.Line (LeftX, Printer.CurrentY + Delta * 3)-(LeftX + grW + Delta, Printer.CurrentY + Delta * 3)
                        Printer.DrawStyle = CurDrawStyle
                     Case "B"
                        Printer.FontBold = True
                     Case "U"
                        Printer.FontUnderline = True
                     Case "C"
                        Printer.ForeColor = Val(Mid(AuxFmt, k))  'lo que queda en el formato es el número del color
                        k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                  End Select
               
                  k = k + 1
               Loop
               
            End If
            
            Printer.CurrentY = Printer.CurrentY + CellHeight * fY
            Printer.CurrentX = LeftX
            X = LeftX
            
            For i = 0 To Grid.Cols - 1
            
               If ColWi(i) <> 0 Then
               
                  If Grid.ColAlignment(i) = FlxAlignCenter Then          'flexAlignCenter
                     Call PrtAlign_(Grid.TextMatrix(GrRow, i), X, ColWi(i), vbCenter)
                  
                  ElseIf Grid.ColAlignment(i) = FlxAlignRight Then      'flexAlignRight
                     Call PrtAlign_(Grid.TextMatrix(GrRow, i), X, ColWi(i), vbRightJustify)
                  
                  Else                                      'flexAlignLeft
                     Call PrtAlign_("  " & Grid.TextMatrix(GrRow, i), X, ColWi(i), vbLeftJustify)
                  End If
                  
                  X = X + ColWi(i)
               
               End If
               
            Next i
         
            GrRow = GrRow + 1
         
            Printer.CurrentY = Printer.CurrentY + CellHeight * fY
            Printer.FontBold = False
            
            If Printer.CurrentY >= Printer.Height - FOOTERSIZE Then
               Exit Do
            End If
         
         Else
            GrRow = GrRow + 1

         End If
         
      Loop

      'imprimimos las líneas verticales
      X = LeftX
      Printer.CurrentY = Printer.CurrentY + CellHeight * fY
      Printer.Line (X, TopY)-(X, Printer.CurrentY + Delta)
      For i = 0 To Grid.Cols - 1
         
         If ColWi(i) <> 0 Then
            X = X + ColWi(i)
            Printer.Line (X + Delta, TopY + AuxTopY(i) * 2.5 * (CellHeight * fY))-(X + Delta, Printer.CurrentY)
         End If
         
      Next i
      
      'imprimimos línea horizontal inferior
      Printer.Line (LeftX, Printer.CurrentY)-(LeftX + grW + Delta, Printer.CurrentY)
      
      If GrRow < Grid.Rows Then
         Call PrtFooter("Continua >>>", RightX)
         Printer.NewPage
         Pag = Pag + 1
      End If
   Loop
   
   CurY = Printer.CurrentY
   Printer.CurrentY = Printer.CurrentY + CellHeight * fY
   
   'imprimimos los totales
   If Total(0) <> "" Then
      If Printer.CurrentY >= Printer.Height - FOOTERSIZE - PRT_LINE_HE * NTotLines Then
         Call PrtFooter("Continua >>>", RightX)
         Printer.NewPage
         Pag = Pag + 1
         
         CurY = Printer.CurrentY
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
      End If
      
      
      For i = 1 To NTotLines
         Call PrtTotales(LeftX, Total, ColWi, Grid, i)
         Printer.CurrentY = Printer.CurrentY + CellHeight * fY * 2
      Next i
   
      'imprimimos línea horizontal inferior
      Printer.Line (LeftX, CurY)-(LeftX, Printer.CurrentY)
      Printer.Line (LeftX + grW + Delta, CurY)-(LeftX + grW + Delta, Printer.CurrentY)
      Printer.Line (LeftX, Printer.CurrentY)-(LeftX + grW + Delta, Printer.CurrentY)
   End If
   
   'imprimimos las observaciones
   If Obs <> "" Then
      If Printer.CurrentY >= Printer.Height - FOOTERSIZE - PRT_LINE_HE * 2 Then
         
         Call PrtFooter("Continua >>>", RightX)
         Printer.NewPage
         Pag = Pag + 1
         
         Printer.Print
         Printer.Print
         Printer.Print
         Printer.Print
      End If

      If Printer.FontSize <= 8 Then
         Printer.FontBold = False
         Printer.FontSize = 9
      End If

      Printer.Print
      
      AuxObs = Obs & " "
      
      Do While AuxObs <> ""
            
         For i = Len(AuxObs) To 1 Step -1
            If Printer.TextWidth(Left(AuxObs, i)) <= grW Then
               j = i
               Do While Mid(AuxObs, j, 1) <> " "
                  j = j - 1
               Loop
               Printer.CurrentX = LeftX
               Printer.Print Left(AuxObs, j)
               Exit For
            End If
         Next i
         
         AuxObs = Mid(AuxObs, j + 1)
      
      Loop
                        
   End If
   
   If CallEndDoc = -1 Then
      Call PrtFooter("Total Págs. " & Pag, RightX)
      Printer.EndDoc
   ElseIf CallEndDoc = -2 Then
      Call PrtFooter("Continua >>>", RightX)
      Printer.EndDoc
   End If
   
   Printer.FontName = OldFName
   Printer.FontBold = OldFBold
   Printer.FontSize = OldFSize
   
   PrtFlexGrid = Pag
   
   On Error GoTo 0
   
End Function
Private Sub PrtTotales(ByVal LeftX As Integer, Total() As String, ColWi() As Integer, Grid As Object, NTotLine As Integer)
   Dim X As Integer
   Dim i As Integer
   Dim j As Integer
   Dim TotBase As Integer
   
   TotBase = (NTotLine - 1) * Grid.Cols

   X = LeftX
   If Total(TotBase) <> "" Then
      ' buscamos primera columna con ancho > 0
      For j = 0 To Grid.Cols - 1
         If ColWi(j) > 0 Then
            If Total(TotBase + j) <> "" Then
               Call PrtAlign_(Total(TotBase), X, ColWi(j), vbCenter)
            Else
               Call PrtAlign_("TOTAL", X, ColWi(j), vbCenter)
            End If
            Exit For
         End If
      Next j
         
      X = X + ColWi(0)
      For i = 1 To Grid.Cols - 1
         If Total(TotBase + i) <> "" Then
            If Grid.ColAlignment(i) = FlxAlignCenter Then          'flexAlignCenter
               Call PrtAlign_(Total(TotBase + i), X, ColWi(i), vbCenter)
            ElseIf Grid.ColAlignment(i) = FlxAlignRight Then      'flexAlignRight
               Call PrtAlign_(Total(TotBase + i), X, ColWi(i), vbRightJustify)
            Else
               Call PrtAlign_(Total(TotBase + i), X, ColWi(i), vbLeftJustify)
            End If
         End If
         X = X + ColWi(i)
      Next i
   End If

End Sub

Private Sub PrtFooter(ByVal StrFooter As String, ByVal RightX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   Printer.CurrentY = Printer.Height - FOOTERPOS
   
   TmpFName = Printer.FontName
   TmpFBold = Printer.FontBold
   TmpFSize = Printer.FontSize
   
   Printer.FontName = FNT_DEFAULT
   Printer.FontBold = False
   Printer.FontSize = 8
   
   Printer.CurrentX = RightX
   Printer.Print StrFooter
   
   Printer.FontName = TmpFName
   Printer.FontBold = TmpFBold
   Printer.FontSize = TmpFSize

End Sub
