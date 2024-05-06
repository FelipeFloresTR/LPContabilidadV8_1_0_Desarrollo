VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrtFlxGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PG_GRCELLHEIGHT = 225                'altura celda por omisión
Private Const PG_LINE_HE = 315                     'alto línea

Private Const DEF_COL_OBLIGATORIA = 1

Private Const FNT_TIMES = "Times New Roman"
Private Const FNT_TITLE = "Arial"
Private Const FNT_COURIER = "Courier New"
Private Const MAX_DEFFNTTIT = 3                   'tamaño arreglo de font de titulos por default

Event EvPrtHeader(PrtPage As Object, LeftX As Long, ByVal RightX As Long, ByVal Pag As Integer)
Public UseCustomHeader As Boolean    'indica si el usuario de la clase imprimirá su propio header

Public PG_FOOTERSIZE As Integer ' = 2000                 '2500
Private PG_FOOTERPOS As Integer ' = PG_FOOTERSIZE - 500

'Imprime la grilla Grid con la siguientes características:
'  Esquina superior izquierda: nombre de la empresa (Nombre)
'  Esquina superior derecha: fecha y numeración paginas
'  Centrado primera línea: Título1
'  Centrado segunda línea: Título2
'  Alineado con Título2: Título3
'  Alineado con Título2: Título4
'  A la izquierda, sobre la grilla: Encabezado
'
'Usa el font de la grilla en pantalla, a menos que UseCourier sea True o se entregue otro Font a través de los campos FontName y FontSize.
'
'Si el alto de la fila es = 0 o el valor asociado a TextMatrix(Fila,ColObligatoria) = "" entonces no imprime la fila.
'
'Ancho de columnas igual a la pantalla, a menos que ColWi(ColObligatoria) sea <> de 0, en cuyo caso usa ColWi(i)
'
'Totales de cada columna se obtienen del arreglo Total()
'
'NTotLines indica la cantidad de lineas de totales que se desea imprimir.
'
'Opcionalmente, puede haber una columna en la grilla con título ".FMT" que contenga:
'  "C(color)" en las filas que se desea imprimir con el color (color). Ejemplo: "C124" imprime la línea con el color 124
'             asegurarse que esta sea la última opción de formato de la línea (ej.: "BC124" imprime la línea en el color seleccionado y con bold
'  "B" en las filas que se desea destacar con bold
'  "L" en las filas en las que se desea insertar una línea divisoria antes de imprimir el registro
'  "M" en las filas en las que se desea insertar una línea divisoria después de imprimir el registro
'  "L(c1,c2)" en las filas en las que se desea insertar una línea divisoria antes de imprimir el registro, entre las columnas c1 y c2
'  "LB" en las filas en las que se desea insertar una línea divisoria antes de imprimir el registro en bold
'  "T" idem a "L" salvo que la línea es entrecortada (pequeños trazos)
'  "FCELL" en las filas que se desea que se tome el formato de la celda: color y bold
'
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
'| Obs                                                       |
'+-----------------------------------------------------------+


Public Grid As Object               'puede ser una MsFlexGrid o una EdFlexGrid (Fairware)

Public Im_Logo As PictureBox        ' Imagen de Logo

Public GrCellHeight As Integer      'altura de las celdas de la grilla

Private lNombres() As String        'arreglo de nombres

Public TabNombres As Integer        'tab para segunda parte de cada nombre, después del vbTab

Private lFntNombres() As FontDef_t   'arreglo de fonts para Nombres()

Private lTitulos() As String         'arreglo de titulos

Private lFntTitulos() As FontDef_t   ' arreglo de fonts para Titulos()

Private lEncabezados() As String     'arreglo de Encabezados()

Public TabEncabezados As Integer     'tab para separar encabezados con cada vbTab en el texto

Private lFntEncabezados() As FontDef_t 'arreglo de fonts para Encabezados()

Public BuffEnc As String           'buffer de texto para poner bajo los encabezados y antes de la grilla

Private lFntBuffEnc As FontDef_t     'font para BuffEnc

Public Obs As String                'bajo al grilla, al lado izquierdo

Public ColObligatoria As Integer    'columna cuyo valor de la celda debe ser distinto de "" para que se imprima la fila,
                                    'si el valor de la celda en esta columna es "", la fila completa no se imprime.
                                    'Si se desea dejar una línea en blanco, esta columna debe tener un " ".
Private lColWi() As Integer         'Arreglo con el ancho de las columnas, sólo válido si ColWi(ColObligatoria) es <> de 0, en cuyo caso usa ColWi(i)

Private lTotal() As String          'arreglo de totales de cada columna.
Public NTotLines As Integer         'NTotLines indica la cantidad de lineas de totales que se desea imprimir.
                                    'Por omisión, se asume 1 línea
                                    'Si NTotLines = 0 no se imprimen totales
                                    'Estas líneas de totales se obtienen del arreglo Total(), suponiendo
                                    'que las filas de totales se alinean a lo largo del arreglo, utilizando
                                    'para cada línea la cantidad de columnas de la grilla. El título de cada
                                    'línea de total se obtiene del mismo arreglo Total(), de acuerdo a las
                                    'columnas en que haya sido ingresado.

Public UseCourier As Integer        'indica si se debe usar Courier en vez del font de la grilla
Public GrFontName As String         'reemplaza el font de la grilla
Public GrFontSize As Integer        'reemplaza el font de la grilla

Public CallEndDoc As Integer        'indica si se debe llamar a la función EndDoc de la impresora, al finalizar
                                    'la impresión. Valores:
                                    '     False (0) no llama a Printer.EndDoc
                                    '     True(-1), valor por omisión, invoca Printer.EndDoc e imprime al final de la última página "Total pags. N"
                                    '     -2, invoca Printer.EndDoc e imprime al final de la última página "Continua >>>"
                                    '     -3, Si tiene más de una franja, invoca Printer.NewPage e imprime al final de la última página "Continua >>>"
                                    '         Si teiene sólo una franja, es como llamarlo con False (0)
                                    'Este campo en Folse es útil cuando se desea
                                    'imprimir algo más después de la grilla, sin cambiar de página.

Public PrintHeader As Boolean       'Indica si se debe imprimir el encabezado (True por omisión) o sólo la grilla (False)

Public EsContinuacion As Boolean    'indica que la impresión es una continuación de otra impresión, por lo que
                                    'no se debe imprimir ningún header ni asignar el CurrentY
                                 
Private lEncabezadosCont() As String 'arreglo de EncabezadosCont(, válido si EsContinuación es True, para poner un encabezado para la continuación.
                                 

Public InitPag As Integer           'número primera página, útil cuando se imprime un doc por partes

Public CellHeight As Integer        'indica la altura de la celda (omisión 225)

Public PrintNumPag As Boolean       'indica si se desea imprimir número de la página actual en encabezado (por omisión True)
Public PrintFecha As Boolean        'indica si se desea imprimir fecha actual en encabezado (por omisión True)
Public PrintHora As Boolean         'indica si se desea imprimir la hora actual en encabezado (por omisión falso)

Public TotFntBold As Boolean           'Entrada: indica si se debe usar bold al imprimir totales

Public PrtDemo As Boolean              'indica si es producto en demo, para que imprima la palabra DEMO en grande

Private FntPageNum As FontDef_t        'almacena la definición del font para número de página y fecha
Private FntGrid  As FontDef_t          'almacena la definición del font de la grilla, tomando en cuenta GrFontName, GrFontSize, UseCourier y el font de la grilla en pantalla

Private DefFntTitulos(MAX_DEFFNTTIT) As FontDef_t  'almacena la definición de font default de los primeros cuatro títulos
Private DefFntNombre As FontDef_t      'almacena la definición de font default de los nombres (igual para todos)
Private DefFntEncabezado As FontDef_t  'almacena la definición de font default para el encabezado
Private DefFntBuffEnc As FontDef_t     'almacena la definición de font default para BuffEnc

Public FmtCol As Integer               'Entrada: indica la línea donde se especifica el formateo (columna con título ".FMT" (se usa cuando la grilla no tiene row de títulos)
Public SplitCol As Integer             'Entrada: indica la columna que debe hacerse split si el texto no cabe en la columna

Public GrLeft As Integer               'Salida: indica margen izquierdo de la grilla (para ser utilizado después de imprimir, si se desea imprimir algo a continuación, manteniendo los márgens de la grilla recién impresa
Public GrRight As Integer              'Salida: indica margen derecho de la grilla (para ser utilizado después de imprimir, si se desea imprimir algo a continuación, manteniendo los márgens de la grilla recién impresa

Public LeftMargin As Integer           'Margen izquierdo mínimo de la hoja, dado que la impresora no es capaz de imprimir tan en el borde izquierdo
Public RightMargin As Integer          'Margen derecho mínimo de la hoja, dado que la impresora no es capaz de imprimir hasta el ancho de la página

'para la impresión de franjas
Private ColIni As Integer, ColFin As Integer, ColPrt As Integer, CntFranja As Integer ' 26 ago 2011: para las franjas

Private Solo1Franja As Boolean         'Indica que sólo se imprimirá 1 franja
Public TieneMasDe1Franja As Boolean    'Indica que la impresión está compuesta por más de una franja
Public PermitirMasDe1Franja As Boolean 'Indica que, a pesar que CallEndDoc = False, permita imprimir más de una franja

Public FixedCols As Integer            'Indica las columnas fijas que se deben imprimir en todas las hojas. Si vale -1, se toman las columnas fijas de la grilla

Private ColsPrtFranja() As Integer     'Columnas que se imprimen en cada franja, considerando las columnas fijas

Private lMembreteCont() As String 'arreglo de MembretesCont(, válido si EsContinuación es True, para poner un encabezado para la continuación.


Public Function PrtFlexGrid(PrtObj As Object) As Integer
   Dim Pag As Integer
   Dim Msg As Boolean
   
   Pag = -1

   If ColIni = -1 Then
      ColIni = 0
   End If
   
   If ColFin = -1 Then
      ColFin = Grid.Cols - 1
   End If
   
   If FixedCols = -1 Then
      FixedCols = Grid.FixedCols
   End If

   If CallEndDoc = 0 And Not PermitirMasDe1Franja Then
      Solo1Franja = True
   End If

   TieneMasDe1Franja = False
   Msg = True
   
   Do While ColPrt < Grid.Cols - 1
      
      
      Pag = PrtFlexGrid1(PrtObj, Msg)
      'Pag = PrtFlexGrid1Membrete(PrtObj, Msg)
      
      If Solo1Franja Or Pag < 0 Then
         Exit Do
      End If
      
      If ColPrt < Grid.Cols - 1 Then
         ColIni = ColPrt + 1
         ColFin = -1
         ColPrt = -1
         CntFranja = CntFranja + 1
      End If
      
      Msg = False
   Loop

   ' Para la siguiente llamada
   ColIni = -1
   ColFin = -1
   ColPrt = -1
   Solo1Franja = False
   FixedCols = -1
   CntFranja = -1

   PrtFlexGrid = Pag
   
End Function
Public Function PrtFlexGrid1(PrtObj As Object, Optional ByVal Msg As Boolean = True) As Integer    'La función Retorna el número de páginas impresas
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   Dim OldFUnderline As Boolean
   Dim OldOrient As Integer
   Dim Linea As String
   Dim Pag As Integer
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim n As Integer
   Dim GrRow As Integer
   Dim TLeft As Integer
   ReDim LinVert(20, 2) As Long
   Dim LeftX As Long, TopY As Long
   Dim RightX As Long
   Dim x As Long
   Dim AuxX As Integer
   Dim fY As Single
   Dim grW As Long
   Dim Delta As Integer
   Dim OfiName As String
   Dim OfiLen As Integer
   Dim AuxFmt As String
   Dim CurDrawStyle As Integer
   Dim AuxObs As String
   Dim TituloCompartido(100) As Integer
   Dim TitWMax As Long
   Dim CurY As Long
   Dim PrtPage As Object
   Dim TopSinNombresCampos As Integer
   Dim LstObs() As String
   Dim X1 As Long, x2 As Long, W1 As Long
   Dim c As Integer, C1 As Integer, C2 As Integer
   Dim AuxCols As String
   Dim GridType As String
   Dim CellSplitHei As Integer
   Dim NSplitLin As Integer
   Dim FmtCell As Boolean
   Dim OldFontBold As Integer
   Dim OldForeColor As Long
   Dim RightColX As Integer
   
   PrtFlexGrid1 = 0
   GridType = LCase(TypeName(Grid))
   
   If Grid.rows <= 1 Then
      MsgBeep vbExclamation
      MsgBox "Listado vacío.", vbExclamation + vbOKOnly
      Exit Function
   End If

   If ColIni = -1 Then
      ColIni = 0
   End If
   
   If ColFin = -1 Then
      ColFin = Grid.Cols - 1
   End If
   
   'ponemos los fonts por default si corresponde
   Call SetDefaultFonts
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   
   Grid.Row = 1
   grW = 0
   
   'armamos el arreglo de columnas que tenemos que imprimir
   ReDim ColsPrtFranja(Grid.Cols)
   
   For n = 0 To Grid.Cols - 1
      ColsPrtFranja(n) = -1
   Next n
  
   If ColIni = 0 Then
      For n = 0 To Grid.Cols - 1
         ColsPrtFranja(n) = n
      Next n
      
   Else
      For n = 0 To FixedCols - 1
         ColsPrtFranja(n) = n
      Next n
      For i = ColIni To Grid.Cols - 1
         ColsPrtFranja(n) = i
         n = n + 1
      Next i
   End If
   
   If UBound(lColWi) = 0 Then
      ReDim lColWi(Grid.Cols)
'      For i = ColIni To ColFin

      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         lColWi(i) = Grid.ColWidth(i)
         If UseCourier = True Then
            lColWi(i) = lColWi(i) * 1.3
         End If
         grW = grW + lColWi(i)
         
      Next n
   Else
'      For i = ColIni To ColFin
         
      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         If UseCourier = True Then
            lColWi(i) = lColWi(i) * 1.3
         End If
         grW = grW + lColWi(i)
         
      Next n
   End If
   
   ColPrt = ColFin
   
   If grW > PrtPage.Width - RightMargin Then
      
      If Msg Then
         If MsgBox1("El listado es más ancho que el tamaño de página seleccionada." & vbCrLf & "Algunas columnas se imprimirán en una hoja aparte." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            PrtFlexGrid1 = -1
            Exit Function
         End If
      End If
      
      If Solo1Franja Then
         MsgBox1 "Este reporte no puede ser impreso en más de una hoja a lo ancho.", vbInformation
      Else
         TieneMasDe1Franja = True
      End If
      
      If FixedCols > 0 Then
         If Msg Then
            If MsgBox1("¿Desea mantener algunas columnas fijas en las páginas laterales?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
               FixedCols = 0
            End If
         End If
      End If
      
      'calculamos el máximo de columnas que podemos imprimir
      grW = 0
      
'      For i = ColIni To ColFin
      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         If grW + lColWi(i) <= PrtPage.Width - RightMargin Then
            grW = grW + lColWi(i)
            ColPrt = i  ' hasta donde llegó
            
            If ColIni = 0 Then
               CntFranja = 1
            End If
         Else
            Exit For
         End If
      Next n
  
   End If
      
   ColFin = ColPrt
   
   'margen izquierdo
   TLeft = 12
   
   'calculamos margen izquierdo de la grilla, considerando que ésta quede centrada
   If CntFranja > 1 Then
      LeftX = LeftMargin + 200
   Else
      LeftX = (PrtPage.Width - 400 - grW) / 2  ' -300 porque a veces no imprime la línea del lado derecho
   End If
   
   If LeftX < 100 Then
      LeftX = LeftMargin
   End If
   
   fY = 0.5
   Delta = 30
   
   'margen derecho para encabezado (N° pág, fecha y continua>>)
   RightX = PrtPage.Width - 2000
   
   'alamcenamos font actual
   OldFName = PrtPage.FontName
   OldFBold = PrtPage.FontBold
   OldFSize = PrtPage.FontSize
   OldFUnderline = PrtPage.FontUnderline
   
   If UseCourier = False Then
      'veamos si hay problema con el font
      On Error Resume Next
      PrtPage.FontName = FNT_TITLE
      PrtPage.FontSize = 10
      PrtPage.FontBold = False
      
      If ERR Then
         MsgBox "Error " & ERR & ", " & Error, vbExclamation
      End If
   End If
   
   For i = 0 To Grid.Cols - 1
      If Trim(Grid.TextMatrix(0, i)) = ".FMT" Then
         FmtCol = i
      End If
   Next i
   
   If InitPag > 0 Then
      Pag = InitPag
   Else
      Pag = 1
   End If
   GrRow = Grid.FixedRows
      
   CellHeight = GrCellHeight
   If CellHeight <= 0 Then
      
      'buscamos la primera fila con altura mayor que cero
      For i = 0 To Grid.rows - 1
         If Grid.RowHeight(i) > 0 Then
            CellHeight = Grid.RowHeight(i)
            Exit For
         End If
      Next i
      
      If CellHeight <= 0 Then
         CellHeight = PG_GRCELLHEIGHT
      End If
   End If
   
   GrLeft = LeftX
   GrRight = LeftX + grW + Delta     'para dibujar líneas
   
   If GrRight > PrtPage.Width - RightMargin Then    '-600 porque si no no imprime hasta el final
      GrRight = PrtPage.Width - RightMargin + Delta
   End If
      
   'Do While GrRow < Grid.Rows
   Do While GrRow < Grid.rows Or (GrRow = Grid.rows And NTotLines > 1)
      
      If Not EsContinuacion Then
               
         'imprimimos encabezado de página
         If PrintHeader = True Then
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, False)
            
         Else
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, True)
            
         End If
         
      ElseIf UBound(lEncabezadosCont) >= 0 Then   'imprimimos encabezados de continuación si hay
         
         If lEncabezadosCont(0) <> "" Then
      
            If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE Then
               Call PrtFooter(PrtPage, "Continua >>>", RightX)
               Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
               Pag = Pag + 1
            
               If PrtPage Is Nothing Then ' no más
                  Exit Do
               End If
   
            End If
            
            For i = 0 To UBound(lEncabezadosCont)
               If lEncabezadosCont(i) <> "" Then
                  Call SetFont(PrtPage, lFntEncabezados(i))
                  Call PrtTxtTabs(PrtPage, lEncabezadosCont(i), False, TabEncabezados, LeftX, RightX)
               End If
            Next i
            
            PrtPage.Print
            
         End If
      
      End If
      
      'definimos font grilla
      If UseCourier = False Then     'si es courier, FntGrid ya tiene la definición correspondiente
         If GrFontName = "" Then
            If GridType = "msflexgrid" Then
               FntGrid.FontName = Grid.FontName
            Else
               FntGrid.FontName = Grid.FlxGrid.FontName
            End If
         Else
            FntGrid.FontName = GrFontName
         End If
         
         If GrFontSize = -1 Then
            If GridType = "msflexgrid" Then
               FntGrid.FontSize = Grid.FontSize
            Else
               FntGrid.FontSize = Grid.FlxGrid.FontSize
            End If
         Else
            FntGrid.FontSize = GrFontSize
         End If
                  
         FntGrid.FontBold = False
      End If
   
      Call SetFont(PrtPage, FntGrid)
   
      TopY = PrtPage.CurrentY
      
      If Grid.FixedRows > 0 Then
         'imprimimos línea horizontal superior
         PrtPage.Print
         PrtPage.Line (LeftX, TopY)-(GrRight - 6, TopY) '1 parche, no imprime la línea a veces sin ese 5
      End If
      
      PrtPage.CurrentX = LeftX
      If Not EsContinuacion Then
         PrtPage.CurrentY = TopY + CellHeight * fY
      End If

      'imprimimos nombres de campos
      PrtPage.FontBold = True
      
      For j = 0 To Grid.FixedRows - 1
         x = LeftX
'         For i = ColIni To ColFin
         For n = 0 To Grid.Cols - 1
            i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
               Exit For
            End If
         
            W1 = lColWi(i)
            'If x + W1 > GrRight Then
'               Exit For
'            End If
         
            If j > 0 Then
               Call PrtAlign_(PrtPage, Trim(Grid.TextMatrix(j, i)), x, W1, vbCenter)
               
            Else
         
               If Grid.FixedAlignment(i) = flexAlignRightCenter Then   'indica que hay un título compartido por dos columnas => línea vertical parte en la siguiente fila
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbRightJustify)
                  TituloCompartido(i) = 1
               ElseIf Grid.FixedAlignment(i) = flexAlignLeftCenter Then
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbLeftJustify)
               Else
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbCenter)
               End If
            
            End If
            
            x = x + lColWi(i)
            
         Next n
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY * 2
                  
      Next j
            
      PrtPage.FontBold = False

      'imprimimos línea horizontal bajo nombres campos
      TopSinNombresCampos = PrtPage.CurrentY + Delta
      If Grid.GridLines <> 0 Then
         PrtPage.Line (LeftX, PrtPage.CurrentY + Delta)-(GrRight, PrtPage.CurrentY + Delta)
      End If
      
      'si ya no hay más registros, generamos una salida del loop, después de haber impreso los nombres de los campos para los totales
      If GrRow = Grid.rows And NTotLines > 1 Then
         GrRow = GrRow + 1
      End If

                  
      'imprimimos registros
      Do While GrRow < Grid.rows
            
         If Grid.TextMatrix(GrRow, ColObligatoria) <> "" And Grid.RowHeight(GrRow) <> 0 Then
            
            'formateo de una línea
            If FmtCol >= 0 Then
               
               AuxFmt = Grid.TextMatrix(GrRow, FmtCol)
               k = 1
               Do While Mid(AuxFmt, k, 1) <> ""
               
                  Select Case Mid(AuxFmt, k, 1)
                     Case "L"       'linea divisoria
                        
                        If Mid(AuxFmt, k + 1, 1) = "(" Then 'vienen las columnas entre las que hay que dibujar la línea (c1,c2)
                           AuxCols = Mid(AuxFmt, k + 1)
                           c = InStr(AuxCols, ",")
                           C1 = 0
                           C2 = 0
                           If c > 0 Then
                              C1 = Val(Mid(AuxCols, 2, c - 2))
                              C2 = Val(Mid(AuxCols, c + 1))
                           End If
                           k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                           
                           X1 = LeftX
                                                      
                           For i = 0 To C1 - 1
                              
                              If lColWi(i) <> 0 Then
                                 X1 = X1 + lColWi(i)
                                 
                                 If X1 > GrRight Then
                                    X1 = GrRight
                                 End If
                                 
                              End If
                              
                           Next i
                           
                           x2 = X1

                           For i = C1 To C2
                              
                              If lColWi(i) <> 0 Then
                                 x2 = x2 + lColWi(i)
                                 
                                 If x2 > GrRight Then
                                    x2 = GrRight
                                 End If
                                 
                              End If
                              
                           Next i
                           
                           PrtPage.Line (X1 + Delta, PrtPage.CurrentY + Delta * 3)-(x2, PrtPage.CurrentY + Delta * 3)
                        Else
                           PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
                        End If
                        
                     Case "T"       'línea de trazos divisoria
                        CurDrawStyle = PrtPage.DrawStyle
                        PrtPage.DrawStyle = vbDash
                        PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
                        PrtPage.DrawStyle = CurDrawStyle
                     Case "B"       'bold
                        PrtPage.FontBold = True
                     Case "U"       'underline
                        PrtPage.FontUnderline = True
                     Case "C"
                        PrtPage.ForeColor = Val(Mid(AuxFmt, k + 1)) 'lo que queda en el formato es el número del color
                        k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                     Case "F"
                        FmtCell = True
                  End Select
               
                  k = k + 1
               Loop
               
            End If
            
            PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
            PrtPage.CurrentX = LeftX
            x = LeftX
            CellSplitHei = 0
            W1 = 0
            
'            For i = ColIni To ColFin
            For n = 0 To Grid.Cols - 1
               i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
                  Exit For
               End If
            
               If lColWi(i) <> 0 Then
               
                  W1 = lColWi(i)
'                  If x + W1 > GrRight Then 'es más ancho que lo que cabe en la página
'                     Exit For
'                  End If

               
                  Grid.Col = i
                  Grid.Row = GrRow
                  
                  
                  If FmtCell Then
                  
                     OldFontBold = PrtPage.FontBold
                     OldForeColor = PrtPage.ForeColor
                     
                     PrtPage.FontBold = Grid.CellFontBold
                     PrtPage.ForeColor = Grid.CellForeColor
                     
                  End If
                  
                  
                  If SplitCol > 0 And i = SplitCol Then
                  
                     NSplitLin = PrtBuf(vbLeftJustify, x + 80, x + lColWi(i) - 80, Grid.TextMatrix(GrRow, i), PrtPage)
                     PrtPage.Print
                     If NSplitLin > 2 Then
                        CellSplitHei = (NSplitLin - 2) * CellHeight * fY
                        PrtPage.CurrentY = PrtPage.CurrentY + CellSplitHei - 100
                     End If
                     
                     PrtPage.CurrentY = PrtPage.CurrentY - (CellHeight - 60)
                 
                  Else
                  
                     If GridType <> "fedgrid" Then
                  
                        If (Grid.ColAlignment(i) = flexAlignCenterCenter And Grid.CellAlignment = 0) Or (Grid.CellAlignment = flexAlignCenterCenter) Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbCenter)
                        
                        ElseIf (Grid.ColAlignment(i) = flexAlignRightCenter And Grid.CellAlignment = 0) Or (Grid.CellAlignment = flexAlignRightCenter) Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbRightJustify)
                        
                        Else                                                  'flexAlignLeftCenter
                           Call PrtAlign_(PrtPage, "  " & Grid.TextMatrix(GrRow, i), x, lColWi(i), vbLeftJustify)
                           
                        End If
                  
                     Else    'FEdGrid primera versión en que no existe Grid.CellAlignment
                     
                        If Grid.ColAlignment(i) = flexAlignCenterCenter Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbCenter)
                        
                        ElseIf Grid.ColAlignment(i) = flexAlignRightCenter Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbRightJustify)
                        
                        Else                                                  'flexAlignLeftCenter
                           Call PrtAlign_(PrtPage, "  " & Grid.TextMatrix(GrRow, i), x, lColWi(i), vbLeftJustify)
                           
                        End If
                     End If
                     
                  End If
                  
                  x = x + lColWi(i)
               
               End If
               
            Next n
         
            If FmtCell Then
                           
               PrtPage.FontBold = OldFontBold
               PrtPage.ForeColor = OldFontBold
               
               FmtCell = False
               
            End If
            
            
            PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
            PrtPage.FontBold = False
            PrtPage.ForeColor = vbBlack
         
            k = InStr(AuxFmt, "M")
            If k Then   'linea divisoria de abajo
                        
               If Mid(AuxFmt, k + 1, 1) = "(" Then 'vienen las columnas entre las que hay que dibujar la línea (c1,c2)
                  AuxCols = Mid(AuxFmt, k + 1)
                  c = InStr(AuxCols, ",")
                  C1 = 0
                  C2 = 0
                  If c > 0 Then
                     C1 = Val(Mid(AuxCols, 2, c - 2))
                     C2 = Val(Mid(AuxCols, c + 1))
                  End If
                  k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                  
                  X1 = LeftX
                                             
                  For i = 0 To C1 - 1
                     
                     If lColWi(i) <> 0 Then
                        X1 = X1 + lColWi(i)
                        
                        If X1 > GrRight Then
                           X1 = GrRight
                        End If
                        
                     End If
                     
                  Next i
                  
                  x2 = X1

                  For i = C1 To C2
                     
                     If lColWi(i) <> 0 Then
                        x2 = x2 + lColWi(i)
                        
                        If x2 > GrRight Then
                           x2 = GrRight
                        End If
                        
                     End If
                     
                  Next i
                  
                  PrtPage.Line (X1 + Delta, PrtPage.CurrentY + Delta * 3)-(x2, PrtPage.CurrentY + Delta * 3)
               Else
                  PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
               End If
               
            End If
         
            GrRow = GrRow + 1
                     
            If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE Then
               Exit Do
            End If
         
         Else
            GrRow = GrRow + 1

         End If
         
      Loop
      
      'imprimimos las líneas verticales
      If Grid.GridLines <> 0 Then
         x = LeftX
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
         
         If Grid.FixedRows = 0 Then
            PrtPage.Line (x, TopSinNombresCampos)-(x, PrtPage.CurrentY + Delta)
         Else
            PrtPage.Line (x, TopY)-(x, PrtPage.CurrentY + Delta)
         End If
         
'         For i = ColIni To ColFin
         For n = 0 To Grid.Cols - 1
            i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
               Exit For
            End If
            
            If lColWi(i) <> 0 Then
               x = x + lColWi(i)
               
'               If x > GrRight Then
'                  Exit For
'               End If
               
               
               If Grid.FixedRows = 0 Then
                  PrtPage.Line (x + Delta, TopSinNombresCampos)-(x + Delta, PrtPage.CurrentY)
               Else
                  PrtPage.Line (x + Delta, TopY + TituloCompartido(i) * 2.5 * (CellHeight * fY))-(x + Delta, PrtPage.CurrentY)
               End If
            End If
            
            
         Next n
         
         RightColX = x + Delta   'posición de la línea vertical más a la derecha
         
      End If
      
      'imprimimos línea horizontal inferior
      If Grid.GridLines <> 0 Then
         PrtPage.Line (LeftX, PrtPage.CurrentY)-(GrRight, PrtPage.CurrentY)
      End If
      
      'If GrRow < Grid.Rows Then
      'si aún no se terminan las líneas o ya se terminaron y falta imprimir más de una linea de totales y no cabe en lo que queda de página, pasamos a página siguiente
      If GrRow < Grid.rows Or (GrRow = Grid.rows And NTotLines > 1 And PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE - PG_LINE_HE * NTotLines) Then
         Dim bHayMas As Boolean, uRow As Integer
         
         uRow = GrRow + 30
         If uRow >= Grid.rows - 1 Then
            uRow = Grid.rows - 1
         End If
         
         bHayMas = False
         For i = GrRow To uRow
            If Grid.TextMatrix(i, ColObligatoria) <> "" Then
               bHayMas = True
               Exit For
            End If
         Next i
      
         If bHayMas = False Then
            Exit Do
         End If
      
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
         Pag = Pag + 1
         
         If PrtPage Is Nothing Then ' no más
            Exit Do
         End If
      
      Else
         Exit Do
            
      End If
   Loop
   
   CurY = PrtPage.CurrentY
   PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
   
   'imprimimos los totales
   If NTotLines > 0 Then
      
      If TotFntBold = True Then
         PrtPage.FontBold = True
      End If
      For i = 1 To NTotLines
         Call PrtTotales(PrtPage, LeftX, lTotal, lColWi, Grid, i)
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY * 2
      Next i
      PrtPage.FontBold = False
      
      'imprimimos línea horizontal inferior y líneas verticales de totales
      PrtPage.Line (LeftX, CurY)-(LeftX, PrtPage.CurrentY)
      'PrtPage.Line (GrRight, CurY)-(GrRight, PrtPage.CurrentY)
      'PrtPage.Line (LeftX, PrtPage.CurrentY)-(GrRight, PrtPage.CurrentY)
      PrtPage.Line (RightColX, CurY)-(RightColX, PrtPage.CurrentY)
      PrtPage.Line (LeftX, PrtPage.CurrentY)-(RightColX, PrtPage.CurrentY)
   End If
   
   'imprimimos las observaciones
   If Obs <> "" And CntFranja <= 1 Then
      If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE - PG_LINE_HE * 3 Then
         
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
         Pag = Pag + 1
         
         If PrtPage Is Nothing Then ' no más
            Exit Function
         End If
         
'         PrtPage.Print
'         PrtPage.Print
'         PrtPage.Print
'         PrtPage.Print

         'imprimimos encabezado de página
         If PrintHeader = True Then
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, False)

         Else
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, True)

         End If

      End If

      PrtPage.FontBold = False
      If PrtPage.FontSize < 9 Then
         PrtPage.FontSize = 9
      End If

      PrtPage.Print
      
      AuxObs = Obs
      k = 0
      i = InStr(AuxObs, vbNewLine)
      Do While i > 0
         
         ReDim Preserve LstObs(k)
         LstObs(k) = Left(AuxObs, i - 1)
         
         AuxObs = Mid(AuxObs, i + Len(vbNewLine))
         i = InStr(AuxObs, vbNewLine)
         k = k + 1
      Loop
      
      If AuxObs <> "" Then   'el último trocito
         ReDim Preserve LstObs(k)
         LstObs(k) = AuxObs
      End If
      
      For k = 0 To UBound(LstObs)
      
         If LstObs(k) = "" Then
            PrtPage.Print
         
         Else
            
            AuxObs = LstObs(k) & " "
            
            Do While AuxObs <> ""
                  
               For i = Len(AuxObs) To 1 Step -1
                  If PrtPage.TextWidth(Left(AuxObs, i)) <= grW Then
                     j = i
                     Do While Mid(AuxObs, j, 1) <> " "
                        j = j - 1
                     Loop
                     PrtPage.CurrentX = LeftX
                     PrtPage.Print Left(AuxObs, j)
                     Exit For
                  End If
               Next i
               
               AuxObs = Mid(AuxObs, j + 1)
            
            Loop
         
         End If
         
      Next k
      
   End If
   
   If CallEndDoc = -1 Then
'      Call PrtFooter(PrtPage, "Total Págs. " & Pag, RightX)
'      PrtPage.EndDoc
      
      If ColFin >= Grid.Cols - 1 Then
         Call PrtFooter(PrtPage, "Total Págs. " & Pag, RightX)
         PrtPage.EndDoc
      Else
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
      End If
      
   ElseIf CallEndDoc = -2 Then    'llama a EndDoc pero pone Continua>>>>
      Call PrtFooter(PrtPage, "Continua >>>", RightX)
      PrtPage.EndDoc
      
   'ElseIf CallEndDoc = -3 Then     '
   ElseIf CallEndDoc = 0 Then     '
   
      If TieneMasDe1Franja Then
         
         If ColFin < Grid.Cols - 1 Then    'no es la última franja
   
            Call PrtFooter(PrtPage, "Continua >>>", RightX)
            Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
            
         End If
         
      End If
      
   End If
   
   PrtPage.FontName = OldFName
   PrtPage.FontBold = OldFBold
   PrtPage.FontSize = OldFSize
   PrtPage.FontUnderline = OldFUnderline
   
   PrtFlexGrid1 = Pag
   
   On Error GoTo 0
   
End Function

Private Sub PrtHeader(PrtPage As Object, LeftX As Long, ByVal TLeft As Long, ByVal RightX As Long, ByVal Pag As Integer, Optional ByVal TinyHeader As Boolean = False)
   Dim i As Integer
   Dim CurX As Integer
   Dim CurY As Integer
   Dim OldFSize As Integer
   Dim OldFName As String
   Dim OldForeColor As Long

   If UseCustomHeader Then
      RaiseEvent EvPrtHeader(PrtPage, LeftX, RightX, Pag)
      Exit Sub
   End If

   PrtPage.CurrentY = 0

   PrtPage.Print
   PrtPage.Print

   If UseCourier = False Then
      PrtPage.FontName = FNT_TITLE
   Else
      PrtPage.FontName = "Courier"
   End If
   
   CurX = PrtPage.CurrentX
   CurY = PrtPage.CurrentY

   If PrtDemo Then
   
      OldFSize = PrtPage.FontSize
      OldFName = PrtPage.FontName
      OldForeColor = PrtPage.ForeColor
      
      PrtPage.FontSize = 96
      PrtPage.FontName = "Arial"
      PrtPage.ForeColor = &HDFDCD7
      
      PrtPage.CurrentX = (PrtPage.Width - PrtPage.TextWidth("DEMO")) / 2
      PrtPage.Print "DEMO"
      
      PrtPage.FontSize = OldFSize
      PrtPage.FontName = OldFName
      PrtPage.ForeColor = OldForeColor

      PrtPage.CurrentX = CurX
      PrtPage.CurrentY = CurY
   End If
   
   PrtPage.FontSize = 10
   PrtPage.Print Tab(TLeft);
   CurX = PrtPage.CurrentX
   If LeftX < CurX Then
      CurX = LeftX
   End If
   
   If LeftX < 0 Then
      LeftX = PrtPage.CurrentX
   End If

   If TinyHeader = False Then
      If Not Im_Logo Is Nothing Then
         If Im_Logo.Picture <> 0 Then
            PrtPage.PaintPicture Im_Logo.Picture, CurX, CurY, Im_Logo.Width, Im_Logo.Height ' No cambia currentX ni CurrentY
         End If
      End If
      
      For i = 0 To UBound(lNombres)
   
         Call SetFont(PrtPage, lFntNombres(i))
         PrtPage.CurrentX = CurX
   
         If i = 0 Then           'número de página
            
            If PrintNumPag Then
               'PrtPage.Print lNombres(i);
               Call PrtTxtTabs(PrtPage, lNombres(i), True, TabNombres, CurX, RightX)
               Call SetFont(PrtPage, FntPageNum)
               PrtPage.CurrentX = RightX
               
               If CntFranja > 0 Then
                  PrtPage.Print "Pág. " & Pag & " - " & CntFranja
               Else
                  PrtPage.Print "Pág. " & Pag
               End If
               
            Else
               'PrtPage.Print lNombres(i)
               Call PrtTxtTabs(PrtPage, lNombres(i), False, TabNombres, CurX, RightX)

            End If
            
         ElseIf i = 1 Then       'fecha
            
            If PrintFecha Then
               'PrtPage.Print lNombres(i);
               Call PrtTxtTabs(PrtPage, lNombres(i), True, TabNombres, CurX, RightX)
               Call SetFont(PrtPage, FntPageNum)
               PrtPage.CurrentX = RightX
               If PrintHora Then
                  PrtPage.Print Format(Now, EDATEFMT & " hh:nn")
               Else
                  PrtPage.Print Format(Now, EDATEFMT)
               End If
            Else
               'PrtPage.Print lNombres(i)
               Call PrtTxtTabs(PrtPage, lNombres(i), False, TabNombres, CurX, RightX)

            End If
            
         ElseIf lNombres(i) <> "" Then
            
            'PrtPage.Print lNombres(i)
            Call PrtTxtTabs(PrtPage, lNombres(i), False, TabNombres, CurX, RightX)

         End If
         
      Next i
      
      PrtPage.Print
   End If
   
   
   For i = 0 To UBound(lTitulos)
      If lTitulos(i) <> "" Then
         Call SetFont(PrtPage, lFntTitulos(i))
         PrtPage.CurrentX = (PrtPage.Width - PrtPage.TextWidth(lTitulos(i))) / 2
         PrtPage.Print lTitulos(i)
      End If
      If TinyHeader = True Then
         Exit For     'sólo imprimimos el primer título
      End If
   Next i

   PrtPage.Print
   'PrtPage.Print

   If TinyHeader = False Then
      For i = 0 To UBound(lEncabezados)
         If lEncabezados(i) <> "" Then
            Call SetFont(PrtPage, lFntEncabezados(i))
            'PrtPage.CurrentX = LeftX
            'PrtPage.Print lEncabezados(i)
            Call PrtTxtTabs(PrtPage, lEncabezados(i), False, TabEncabezados, LeftX, RightX)
         End If
      Next i
      
      PrtPage.Print
      
      If BuffEnc <> "" Then
                  
         Call SetFont(PrtPage, lFntBuffEnc)
         
         PrtPage.Print
         PrtPage.Print
         
         PrtPage.CurrentX = LeftX
   
         Call PrtBuf(vbLeftJustify, LeftX, RightX, BuffEnc, PrtPage)
         
         PrtPage.Print
         PrtPage.Print
      
      End If
      
   End If
   
End Sub

Public Sub PrtAlign_(PrtPage As Object, ByVal Txt As String, ByVal x As Integer, ByVal W As Integer, ByVal Align As Integer)
   Dim Tw As Integer
   Dim Aux As String
   

   Aux = Txt
   Aux = ReplaceStr(Aux, vbCrLf, " ")
   Aux = ReplaceStr(Aux, vbLf, " ")
   Aux = ReplaceStr(Aux, vbCr, " ")
   
   Tw = PrtPage.TextWidth(Aux)
   
   Do While Tw > W
      Aux = Left(Aux, Len(Aux) - 1)
      Tw = PrtPage.TextWidth(Aux)
   Loop

   Select Case Align
      Case vbLeftJustify:
         PrtPage.CurrentX = x
      Case vbRightJustify:
         PrtPage.CurrentX = x + (W - Tw)
      Case vbCenter:
         PrtPage.CurrentX = x + (W - Tw) / 2
   End Select

   PrtPage.Print Aux;

End Sub

Public Sub PrtTxtTabs(PrtPage As Object, ByVal Txt As String, ByVal SemiCol As Boolean, ByVal TabTxt As Integer, ByVal CurX As Integer, ByVal RightX As Integer)
   Dim Txt1 As String, Txt2 As String
   Dim i As Integer, j As Integer
   
   i = InStr(Txt, vbTab)
   j = InStr(Txt, vbCr & vbLf)
   PrtPage.CurrentX = CurX
   
   If i <> 0 And TabTxt <> 0 Then
      
      Txt1 = Left(Txt, i - 1)
      Txt2 = Mid(Txt, i + 1)
      
      PrtPage.Print Txt1;
      PrtPage.CurrentX = CurX + TabTxt
      
      If SemiCol Then
         PrtPage.Print Txt2;
      Else
         PrtPage.Print Txt2
      End If
      
   ElseIf j <> 0 Then
      
      Call PrtBuf(vbLeftJustify, CurX, RightX, Txt, PrtPage)
   
   Else
   
      If SemiCol Then
         PrtPage.Print Txt;
      Else
         'PrtPage.Print Txt
         Call PrtBuf(vbLeftJustify, CurX, RightX, Txt, PrtPage)
         PrtPage.Print
      End If
      
End If

End Sub

Public Sub PrtFooter(PrtPage As Object, ByVal StrFooter As String, ByVal RightX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   PrtPage.CurrentY = PrtPage.Height - PG_FOOTERPOS
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 8
   
   PrtPage.CurrentX = RightX
   PrtPage.Print StrFooter
   
   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

Private Sub PrtTotales(PrtPage As Object, ByVal LeftX As Integer, Total() As String, ColWi() As Integer, Grid As Object, NTotLine As Integer)
   Dim x As Integer
   Dim i As Integer
   Dim j As Integer
   Dim n As Integer
   Dim TotBase As Integer
   
   TotBase = (NTotLine - 1) * Grid.Cols

   x = LeftX
   
'   For i = ColIni To ColFin
   For n = 0 To Grid.Cols - 1
      i = ColsPrtFranja(n)
      If i < 0 Or i > ColFin Then
         Exit For
      End If
      If Total(TotBase + i) <> "" And ColWi(i) > 0 Then
         If Grid.ColAlignment(i) = flexAlignCenterCenter Then
            Call PrtAlign_(PrtPage, Total(TotBase + i), x, ColWi(i), vbCenter)
         ElseIf Grid.ColAlignment(i) = flexAlignRightCenter Then
            Call PrtAlign_(PrtPage, Total(TotBase + i), x, ColWi(i), vbRightJustify)
         Else                          'flexAlignLeftCenter
            Call PrtAlign_(PrtPage, Total(TotBase + i), x, ColWi(i), vbLeftJustify)
         End If
      End If
      x = x + ColWi(i)
   Next n

End Sub

Public Sub Clear()
               
   ReDim lColWi(0)
   lColWi(0) = 0
   
   ReDim lTotal(0)
   lTotal(0) = ""
   NTotLines = 1
   
   UseCourier = False
   ColObligatoria = 1
   CallEndDoc = -1
   PrintHeader = True
   EsContinuacion = False
   LeftMargin = 100
   RightMargin = 600
   
   GrFontName = ""
   GrFontSize = -1
   
   Obs = ""
   InitPag = -1
   CellHeight = 0
   
   FmtCol = -1
   
   PrintNumPag = True
   PrintFecha = True
   PrintHora = False
   
   'definimos font de número de página
   FntPageNum.FontName = FNT_TITLE
   FntPageNum.FontSize = 8
   FntPageNum.FontBold = False
   FntPageNum.FontUnderline = False
   
   'font default para grilla
   FntGrid.FontName = FNT_COURIER
   FntGrid.FontSize = 10
   FntGrid.FontBold = False
   FntGrid.FontUnderline = False
   
   'font default para Nombres
   DefFntNombre.FontName = "Arial"
   DefFntNombre.FontSize = 9
   DefFntNombre.FontBold = False
   DefFntNombre.FontUnderline = False
   
   'tab para segunda parte de nombre, después de vbTab (opcional)
   TabNombres = 0
   
   'font default para Encabezado
   DefFntEncabezado.FontName = "Arial"
   DefFntEncabezado.FontSize = 10
   DefFntEncabezado.FontBold = False
   DefFntEncabezado.FontUnderline = False
   
   'font default para BuffEnc
   DefFntBuffEnc.FontName = "Arial"
   DefFntBuffEnc.FontSize = 8
   DefFntBuffEnc.FontBold = False
   DefFntBuffEnc.FontUnderline = False
   
   'font default para titulos
   DefFntTitulos(0).FontName = "Arial"
   DefFntTitulos(0).FontSize = 14
   DefFntTitulos(0).FontBold = True
   DefFntTitulos(0).FontUnderline = False
   
   DefFntTitulos(1).FontName = "Arial"
   DefFntTitulos(1).FontSize = 12
   DefFntTitulos(1).FontBold = False
   DefFntTitulos(1).FontUnderline = False
   
   DefFntTitulos(2).FontName = "Arial"
   DefFntTitulos(2).FontSize = 10
   DefFntTitulos(2).FontBold = False
   DefFntTitulos(2).FontUnderline = False
   
   DefFntTitulos(3).FontName = "Arial"
   DefFntTitulos(3).FontSize = 10
   DefFntTitulos(3).FontBold = False
   DefFntTitulos(3).FontUnderline = False
   
   TotFntBold = True
   PrtDemo = False
   
   ColIni = -1
   ColFin = -1
   ColPrt = -1
   CntFranja = -1
   Solo1Franja = False
   TieneMasDe1Franja = False
   PermitirMasDe1Franja = False
   FixedCols = -1
   
   ReDim ColsPrtFranja(0)
   ColsPrtFranja(0) = 0


End Sub

Private Function SetFont(PrtPage As Object, Fnt As FontDef_t)
   PrtPage.FontName = Fnt.FontName
   PrtPage.FontSize = Fnt.FontSize
   PrtPage.FontBold = Fnt.FontBold
   PrtPage.FontUnderline = Fnt.FontUnderline
End Function

Public Property Get Nombres() As Variant

End Property

Public Property Let Nombres(ByVal vNewValue As Variant)
   Dim i As Integer

   ReDim lNombres(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lNombres(i) = vNewValue(i)
   Next i

End Property

Public Property Get Titulos() As Variant

End Property

Public Property Let Titulos(ByVal vNewValue As Variant)
   Dim i As Integer
   Dim Idx As Integer

   ReDim lTitulos(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lTitulos(i) = vNewValue(i)
   Next i
   
End Property

Public Property Get Encabezados() As Variant

End Property

Public Property Let Encabezados(ByVal vNewValue As Variant)
   Dim i As Integer

   ReDim lEncabezados(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lEncabezados(i) = vNewValue(i)
   Next i
   
End Property
Public Property Get EncabezadosCont() As Variant

End Property

Public Property Let EncabezadosCont(ByVal vNewValue As Variant)
   Dim i As Integer

   ReDim lEncabezadosCont(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lEncabezadosCont(i) = vNewValue(i)
   Next i
   
End Property
Private Sub Class_Initialize()
   
   ReDim lColWi(0)
   ReDim lNombres(0)
   ReDim lFntNombres(0)
   ReDim lTitulos(0)
   ReDim lFntTitulos(0)
   ReDim lEncabezados(0)
   ReDim lEncabezadosCont(0)
   ReDim lFntEncabezados(0)
   
   FOOTERSIZE = 2000

   Call Clear

End Sub

Public Property Let ColWi(ByVal vNewValue As Variant)
   Dim i As Integer

   ReDim lColWi(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lColWi(i) = vNewValue(i)
   Next i

End Property
Public Property Get Total() As Variant

End Property

Public Property Let Total(ByVal vNewValue As Variant)
   Dim i As Integer

   ReDim lTotal(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lTotal(i) = vNewValue(i)
   Next i

End Property

Friend Sub FntTitulos(vNewValue() As FontDef_t)
   Dim i As Integer

   ReDim lFntTitulos(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lFntTitulos(i).FontName = vNewValue(i).FontName
      lFntTitulos(i).FontSize = vNewValue(i).FontSize
      lFntTitulos(i).FontBold = vNewValue(i).FontBold
      lFntTitulos(i).FontUnderline = vNewValue(i).FontUnderline
   Next i

End Sub

Friend Sub FntNombres(vNewValue() As FontDef_t)
   Dim i As Integer

   ReDim lFntNombres(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lFntNombres(i).FontName = vNewValue(i).FontName
      lFntNombres(i).FontSize = vNewValue(i).FontSize
      lFntNombres(i).FontBold = vNewValue(i).FontBold
      lFntNombres(i).FontUnderline = vNewValue(i).FontUnderline
   Next i

End Sub

Friend Sub FntEncabezados(vNewValue() As FontDef_t)
   Dim i As Integer

   ReDim lFntEncabezados(UBound(vNewValue))
   For i = 0 To UBound(vNewValue)
      lFntEncabezados(i).FontName = vNewValue(i).FontName
      lFntEncabezados(i).FontSize = vNewValue(i).FontSize
      lFntEncabezados(i).FontBold = vNewValue(i).FontBold
      lFntEncabezados(i).FontUnderline = vNewValue(i).FontUnderline
   Next i

End Sub

Friend Sub FntBuffEnc(vNewValue As FontDef_t)

   lFntBuffEnc.FontName = vNewValue.FontName
   lFntBuffEnc.FontSize = vNewValue.FontSize
   lFntBuffEnc.FontBold = vNewValue.FontBold
   lFntBuffEnc.FontUnderline = vNewValue.FontUnderline

End Sub

Private Function SetDefaultFonts()
   Dim i As Integer
   Dim Idx As Integer
   
   ReDim Preserve lFntNombres(UBound(lNombres))
   
   If lFntNombres(0).FontName = "" Then
      lFntNombres(0).FontName = DefFntNombre.FontName
      lFntNombres(0).FontSize = DefFntNombre.FontSize
      lFntNombres(0).FontBold = DefFntNombre.FontBold
      lFntNombres(0).FontUnderline = DefFntNombre.FontUnderline
   End If
   
   For i = 1 To UBound(lFntNombres)
      If lFntNombres(i).FontName = "" Then
         lFntNombres(i).FontName = lFntNombres(0).FontName
         lFntNombres(i).FontSize = lFntNombres(0).FontSize
         lFntNombres(i).FontBold = lFntNombres(0).FontBold
         lFntNombres(i).FontUnderline = lFntNombres(0).FontUnderline
      End If
   Next i
      
   ReDim Preserve lFntTitulos(UBound(lTitulos))

   For i = 0 To UBound(lFntTitulos)
      If lFntTitulos(i).FontName = "" Then
         If i <= MAX_DEFFNTTIT Then
            Idx = i
         Else
            Idx = MAX_DEFFNTTIT
         End If
         
         lFntTitulos(i).FontName = DefFntTitulos(Idx).FontName
         lFntTitulos(i).FontSize = DefFntTitulos(Idx).FontSize
         lFntTitulos(i).FontBold = DefFntTitulos(Idx).FontBold
         lFntTitulos(i).FontUnderline = DefFntTitulos(Idx).FontUnderline
      End If
   Next i

   ReDim Preserve lFntEncabezados(UBound(lEncabezados))
   
   If lFntEncabezados(0).FontName = "" Then
      lFntEncabezados(0).FontName = DefFntEncabezado.FontName
      lFntEncabezados(0).FontSize = DefFntEncabezado.FontSize
      lFntEncabezados(0).FontBold = DefFntEncabezado.FontBold
      lFntEncabezados(0).FontUnderline = DefFntEncabezado.FontUnderline
   End If
   
   For i = 1 To UBound(lFntEncabezados)
      If lFntEncabezados(i).FontName = "" Then
         lFntEncabezados(i).FontName = lFntEncabezados(0).FontName
         lFntEncabezados(i).FontSize = lFntEncabezados(0).FontSize
         lFntEncabezados(i).FontBold = lFntEncabezados(0).FontBold
         lFntEncabezados(i).FontUnderline = lFntEncabezados(0).FontUnderline
      End If
   Next i

   If lFntBuffEnc.FontName = "" Then
      lFntBuffEnc.FontName = DefFntBuffEnc.FontName
      lFntBuffEnc.FontSize = DefFntBuffEnc.FontSize
      lFntBuffEnc.FontBold = DefFntBuffEnc.FontBold
      lFntBuffEnc.FontUnderline = DefFntBuffEnc.FontUnderline
   End If
   
End Function

Public Sub SetSmallFont(Optional ByVal FntSize As Integer = 8)
   
   Me.GrFontName = "Arial"
   Me.GrFontSize = FntSize

End Sub

Public Sub UndoSetSmallFont()
   
   Me.GrFontName = ""
   Me.GrFontSize = -1

End Sub

Public Property Get FOOTERSIZE() As Integer
   FOOTERSIZE = PG_FOOTERSIZE
End Property

Public Property Let FOOTERSIZE(ByVal Value As Integer)

   PG_FOOTERSIZE = Value
   
   If PG_FOOTERSIZE > 500 Then
      PG_FOOTERPOS = PG_FOOTERSIZE - 500
   Else
      PG_FOOTERPOS = 0
   End If

End Property


'pipe 2860036
Public Function PrtFlexGrid1Membrete(PrtObj As Object, Optional ByVal Msg As Boolean = True) As Integer    'La función Retorna el número de páginas impresas
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   Dim OldFUnderline As Boolean
   Dim OldOrient As Integer
   Dim Linea As String
   Dim Pag As Integer
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim n As Integer
   Dim GrRow As Integer
   Dim TLeft As Integer
   ReDim LinVert(20, 2) As Long
   Dim LeftX As Long, TopY As Long
   Dim RightX As Long
   Dim x As Long
   Dim AuxX As Integer
   Dim fY As Single
   Dim grW As Long
   Dim Delta As Integer
   Dim OfiName As String
   Dim OfiLen As Integer
   Dim AuxFmt As String
   Dim CurDrawStyle As Integer
   Dim AuxObs As String
   Dim TituloCompartido(100) As Integer
   Dim TitWMax As Long
   Dim CurY As Long
   Dim PrtPage As Object
   Dim TopSinNombresCampos As Integer
   Dim LstObs() As String
   Dim X1 As Long, x2 As Long, W1 As Long
   Dim c As Integer, C1 As Integer, C2 As Integer
   Dim AuxCols As String
   Dim GridType As String
   Dim CellSplitHei As Integer
   Dim NSplitLin As Integer
   Dim FmtCell As Boolean
   Dim OldFontBold As Integer
   Dim OldForeColor As Long
   Dim RightColX As Integer
   
   PrtFlexGrid1Membrete = 0
   GridType = LCase(TypeName(Grid))
   
   If Grid.rows <= 1 Then
      MsgBeep vbExclamation
      MsgBox "Listado vacío.", vbExclamation + vbOKOnly
      Exit Function
   End If

   If ColIni = -1 Then
      ColIni = 0
   End If
   
   If ColFin = -1 Then
      ColFin = Grid.Cols - 1
   End If
   
   'ponemos los fonts por default si corresponde
   Call SetDefaultFonts
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
   
   Grid.Row = 1
   grW = 0
   
   'armamos el arreglo de columnas que tenemos que imprimir
   ReDim ColsPrtFranja(Grid.Cols)
   
   For n = 0 To Grid.Cols - 1
      ColsPrtFranja(n) = -1
   Next n
  
   If ColIni = 0 Then
      For n = 0 To Grid.Cols - 1
         ColsPrtFranja(n) = n
      Next n
      
   Else
      For n = 0 To FixedCols - 1
         ColsPrtFranja(n) = n
      Next n
      For i = ColIni To Grid.Cols - 1
         ColsPrtFranja(n) = i
         n = n + 1
      Next i
   End If
   
   If UBound(lColWi) = 0 Then
      ReDim lColWi(Grid.Cols)
'      For i = ColIni To ColFin

      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         lColWi(i) = Grid.ColWidth(i)
         If UseCourier = True Then
            lColWi(i) = lColWi(i) * 1.3
         End If
         grW = grW + lColWi(i)
         
      Next n
   Else
'      For i = ColIni To ColFin
         
      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         If UseCourier = True Then
            lColWi(i) = lColWi(i) * 1.3
         End If
         grW = grW + lColWi(i)
         
      Next n
   End If
   
   ColPrt = ColFin
   
   If grW > PrtPage.Width - RightMargin Then
      
      If Msg Then
         If MsgBox1("El listado es más ancho que el tamaño de página seleccionada." & vbCrLf & "Algunas columnas se imprimirán en una hoja aparte." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            PrtFlexGrid1Membrete = -1
            Exit Function
         End If
      End If
      
      If Solo1Franja Then
         MsgBox1 "Este reporte no puede ser impreso en más de una hoja a lo ancho.", vbInformation
      Else
         TieneMasDe1Franja = True
      End If
      
      If FixedCols > 0 Then
         If Msg Then
            If MsgBox1("¿Desea mantener algunas columnas fijas en las páginas laterales?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
               FixedCols = 0
            End If
         End If
      End If
      
      'calculamos el máximo de columnas que podemos imprimir
      grW = 0
      
'      For i = ColIni To ColFin
      For n = 0 To Grid.Cols - 1
         i = ColsPrtFranja(n)
         If i < 0 Then
            Exit For
         End If
         
         If grW + lColWi(i) <= PrtPage.Width - RightMargin Then
            grW = grW + lColWi(i)
            ColPrt = i  ' hasta donde llegó
            
            If ColIni = 0 Then
               CntFranja = 1
            End If
         Else
            Exit For
         End If
      Next n
  
   End If
      
   ColFin = ColPrt
   
   'margen izquierdo
   TLeft = 12
   
   'calculamos margen izquierdo de la grilla, considerando que ésta quede centrada
   If CntFranja > 1 Then
      LeftX = LeftMargin + 200
   Else
      LeftX = (PrtPage.Width - 400 - grW) / 2  ' -300 porque a veces no imprime la línea del lado derecho
   End If
   
   If LeftX < 100 Then
      LeftX = LeftMargin
   End If
   
   fY = 0.5
   Delta = 30
   
   'margen derecho para encabezado (N° pág, fecha y continua>>)
   RightX = PrtPage.Width - 2000
   
   'alamcenamos font actual
   OldFName = PrtPage.FontName
   OldFBold = PrtPage.FontBold
   OldFSize = PrtPage.FontSize
   OldFUnderline = PrtPage.FontUnderline
   
   If UseCourier = False Then
      'veamos si hay problema con el font
      On Error Resume Next
      PrtPage.FontName = FNT_TITLE
      PrtPage.FontSize = 10
      PrtPage.FontBold = False
      
      If ERR Then
         MsgBox "Error " & ERR & ", " & Error, vbExclamation
      End If
   End If
   
   For i = 0 To Grid.Cols - 1
      If Trim(Grid.TextMatrix(0, i)) = ".FMT" Then
         FmtCol = i
      End If
   Next i
   
   If InitPag > 0 Then
      Pag = InitPag
   Else
      Pag = 1
   End If
   GrRow = Grid.FixedRows
      
   CellHeight = GrCellHeight
   If CellHeight <= 0 Then
      
      'buscamos la primera fila con altura mayor que cero
      For i = 0 To Grid.rows - 1
         If Grid.RowHeight(i) > 0 Then
            CellHeight = Grid.RowHeight(i)
            Exit For
         End If
      Next i
      
      If CellHeight <= 0 Then
         CellHeight = PG_GRCELLHEIGHT
      End If
   End If
   
   GrLeft = LeftX
   GrRight = LeftX + grW + Delta     'para dibujar líneas
   
   If GrRight > PrtPage.Width - RightMargin Then    '-600 porque si no no imprime hasta el final
      GrRight = PrtPage.Width - RightMargin + Delta
   End If
      
   'Do While GrRow < Grid.Rows
   Do While GrRow < Grid.rows Or (GrRow = Grid.rows And NTotLines > 1)
      
      If Not EsContinuacion Then
               
         'imprimimos encabezado de página
         If PrintHeader = True Then
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, False)
            
         Else
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, True)
            
         End If
         
      ElseIf UBound(lEncabezadosCont) >= 0 Then   'imprimimos encabezados de continuación si hay
         
         If lEncabezadosCont(0) <> "" Then
      
            If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE Then
               Call PrtFooter(PrtPage, "Continua >>>", RightX)
               Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
               Pag = Pag + 1
            
               If PrtPage Is Nothing Then ' no más
                  Exit Do
               End If
   
            End If
            
            For i = 0 To UBound(lEncabezadosCont)
               If lEncabezadosCont(i) <> "" Then
                  Call SetFont(PrtPage, lFntEncabezados(i))
                  Call PrtTxtTabs(PrtPage, lEncabezadosCont(i), False, TabEncabezados, LeftX, RightX)
               End If
            Next i
            
            PrtPage.Print
            
         End If
      
      End If
      
      
      'definimos font grilla
      If UseCourier = False Then     'si es courier, FntGrid ya tiene la definición correspondiente
         If GrFontName = "" Then
            If GridType = "msflexgrid" Then
               FntGrid.FontName = Grid.FontName
            Else
               FntGrid.FontName = Grid.FlxGrid.FontName
            End If
         Else
            FntGrid.FontName = GrFontName
         End If
         
         If GrFontSize = -1 Then
            If GridType = "msflexgrid" Then
               FntGrid.FontSize = Grid.FontSize
            Else
               FntGrid.FontSize = Grid.FlxGrid.FontSize
            End If
         Else
            FntGrid.FontSize = GrFontSize
         End If
                  
         FntGrid.FontBold = False
      End If
   
      Call SetFont(PrtPage, FntGrid)
   
      TopY = PrtPage.CurrentY
      
      If Grid.FixedRows > 0 Then
         'imprimimos línea horizontal superior
         PrtPage.Print
         PrtPage.Line (LeftX, TopY)-(GrRight - 6, TopY) '1 parche, no imprime la línea a veces sin ese 5
      End If
      
      PrtPage.CurrentX = LeftX
      If Not EsContinuacion Then
         PrtPage.CurrentY = TopY + CellHeight * fY
      End If

      'imprimimos nombres de campos
      PrtPage.FontBold = True
      
      For j = 0 To Grid.FixedRows - 1
         x = LeftX
'         For i = ColIni To ColFin
         For n = 0 To Grid.Cols - 1
            i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
               Exit For
            End If
         
            W1 = lColWi(i)
            'If x + W1 > GrRight Then
'               Exit For
'            End If
         
            If j > 0 Then
               Call PrtAlign_(PrtPage, Trim(Grid.TextMatrix(j, i)), x, W1, vbCenter)
               
            Else
         
               If Grid.FixedAlignment(i) = flexAlignRightCenter Then   'indica que hay un título compartido por dos columnas => línea vertical parte en la siguiente fila
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbRightJustify)
                  TituloCompartido(i) = 1
               ElseIf Grid.FixedAlignment(i) = flexAlignLeftCenter Then
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbLeftJustify)
               Else
                  Call PrtAlign_(PrtPage, Grid.TextMatrix(j, i), x, W1, vbCenter)
               End If
            
            End If
            
            x = x + lColWi(i)
            
         Next n
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY * 2
                  
      Next j
            
      PrtPage.FontBold = False

      'imprimimos línea horizontal bajo nombres campos
      TopSinNombresCampos = PrtPage.CurrentY + Delta
      If Grid.GridLines <> 0 Then
         PrtPage.Line (LeftX, PrtPage.CurrentY + Delta)-(GrRight, PrtPage.CurrentY + Delta)
      End If
      
      'si ya no hay más registros, generamos una salida del loop, después de haber impreso los nombres de los campos para los totales
      If GrRow = Grid.rows And NTotLines > 1 Then
         GrRow = GrRow + 1
      End If

                  
      'imprimimos registros
      Do While GrRow < Grid.rows
            
         If Grid.TextMatrix(GrRow, ColObligatoria) <> "" And Grid.RowHeight(GrRow) <> 0 Then
            
            'formateo de una línea
            If FmtCol >= 0 Then
               
               AuxFmt = Grid.TextMatrix(GrRow, FmtCol)
               k = 1
               Do While Mid(AuxFmt, k, 1) <> ""
               
                  Select Case Mid(AuxFmt, k, 1)
                     Case "L"       'linea divisoria
                        
                        If Mid(AuxFmt, k + 1, 1) = "(" Then 'vienen las columnas entre las que hay que dibujar la línea (c1,c2)
                           AuxCols = Mid(AuxFmt, k + 1)
                           c = InStr(AuxCols, ",")
                           C1 = 0
                           C2 = 0
                           If c > 0 Then
                              C1 = Val(Mid(AuxCols, 2, c - 2))
                              C2 = Val(Mid(AuxCols, c + 1))
                           End If
                           k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                           
                           X1 = LeftX
                                                      
                           For i = 0 To C1 - 1
                              
                              If lColWi(i) <> 0 Then
                                 X1 = X1 + lColWi(i)
                                 
                                 If X1 > GrRight Then
                                    X1 = GrRight
                                 End If
                                 
                              End If
                              
                           Next i
                           
                           x2 = X1

                           For i = C1 To C2
                              
                              If lColWi(i) <> 0 Then
                                 x2 = x2 + lColWi(i)
                                 
                                 If x2 > GrRight Then
                                    x2 = GrRight
                                 End If
                                 
                              End If
                              
                           Next i
                           
                           PrtPage.Line (X1 + Delta, PrtPage.CurrentY + Delta * 3)-(x2, PrtPage.CurrentY + Delta * 3)
                        Else
                           PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
                        End If
                        
                     Case "T"       'línea de trazos divisoria
                        CurDrawStyle = PrtPage.DrawStyle
                        PrtPage.DrawStyle = vbDash
                        PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
                        PrtPage.DrawStyle = CurDrawStyle
                     Case "B"       'bold
                        PrtPage.FontBold = True
                     Case "U"       'underline
                        PrtPage.FontUnderline = True
                     Case "C"
                        PrtPage.ForeColor = Val(Mid(AuxFmt, k + 1)) 'lo que queda en el formato es el número del color
                        k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                     Case "F"
                        FmtCell = True
                  End Select
               
                  k = k + 1
               Loop
               
            End If
            
            PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
            PrtPage.CurrentX = LeftX
            x = LeftX
            CellSplitHei = 0
            W1 = 0
            
'            For i = ColIni To ColFin
            For n = 0 To Grid.Cols - 1
               i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
                  Exit For
               End If
            
               If lColWi(i) <> 0 Then
               
                  W1 = lColWi(i)
'                  If x + W1 > GrRight Then 'es más ancho que lo que cabe en la página
'                     Exit For
'                  End If

               
                  Grid.Col = i
                  Grid.Row = GrRow
                  
                  
                  If FmtCell Then
                  
                     OldFontBold = PrtPage.FontBold
                     OldForeColor = PrtPage.ForeColor
                     
                     PrtPage.FontBold = Grid.CellFontBold
                     PrtPage.ForeColor = Grid.CellForeColor
                     
                  End If
                  
                  
                  If SplitCol > 0 And i = SplitCol Then
                  
                     NSplitLin = PrtBuf(vbLeftJustify, x + 80, x + lColWi(i) - 80, Grid.TextMatrix(GrRow, i), PrtPage)
                     PrtPage.Print
                     If NSplitLin > 2 Then
                        CellSplitHei = (NSplitLin - 2) * CellHeight * fY
                        PrtPage.CurrentY = PrtPage.CurrentY + CellSplitHei - 100
                     End If
                     
                     PrtPage.CurrentY = PrtPage.CurrentY - (CellHeight - 60)
                 
                  Else
                  
                     If GridType <> "fedgrid" Then
                  
                        If (Grid.ColAlignment(i) = flexAlignCenterCenter And Grid.CellAlignment = 0) Or (Grid.CellAlignment = flexAlignCenterCenter) Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbCenter)
                        
                        ElseIf (Grid.ColAlignment(i) = flexAlignRightCenter And Grid.CellAlignment = 0) Or (Grid.CellAlignment = flexAlignRightCenter) Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbRightJustify)
                        
                        Else                                                  'flexAlignLeftCenter
                           Call PrtAlign_(PrtPage, "  " & Grid.TextMatrix(GrRow, i), x, lColWi(i), vbLeftJustify)
                           
                        End If
                  
                     Else    'FEdGrid primera versión en que no existe Grid.CellAlignment
                     
                        If Grid.ColAlignment(i) = flexAlignCenterCenter Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbCenter)
                        
                        ElseIf Grid.ColAlignment(i) = flexAlignRightCenter Then
                           Call PrtAlign_(PrtPage, Grid.TextMatrix(GrRow, i), x, lColWi(i), vbRightJustify)
                        
                        Else                                                  'flexAlignLeftCenter
                           Call PrtAlign_(PrtPage, "  " & Grid.TextMatrix(GrRow, i), x, lColWi(i), vbLeftJustify)
                           
                        End If
                     End If
                     
                  End If
                  
                  x = x + lColWi(i)
               
               End If
               
            Next n
         
            If FmtCell Then
                           
               PrtPage.FontBold = OldFontBold
               PrtPage.ForeColor = OldFontBold
               
               FmtCell = False
               
            End If
            
            
            PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
            PrtPage.FontBold = False
            PrtPage.ForeColor = vbBlack
         
            k = InStr(AuxFmt, "M")
            If k Then   'linea divisoria de abajo
                        
               If Mid(AuxFmt, k + 1, 1) = "(" Then 'vienen las columnas entre las que hay que dibujar la línea (c1,c2)
                  AuxCols = Mid(AuxFmt, k + 1)
                  c = InStr(AuxCols, ",")
                  C1 = 0
                  C2 = 0
                  If c > 0 Then
                     C1 = Val(Mid(AuxCols, 2, c - 2))
                     C2 = Val(Mid(AuxCols, c + 1))
                  End If
                  k = Len(Grid.TextMatrix(GrRow, FmtCol)) + 1
                  
                  X1 = LeftX
                                             
                  For i = 0 To C1 - 1
                     
                     If lColWi(i) <> 0 Then
                        X1 = X1 + lColWi(i)
                        
                        If X1 > GrRight Then
                           X1 = GrRight
                        End If
                        
                     End If
                     
                  Next i
                  
                  x2 = X1

                  For i = C1 To C2
                     
                     If lColWi(i) <> 0 Then
                        x2 = x2 + lColWi(i)
                        
                        If x2 > GrRight Then
                           x2 = GrRight
                        End If
                        
                     End If
                     
                  Next i
                  
                  PrtPage.Line (X1 + Delta, PrtPage.CurrentY + Delta * 3)-(x2, PrtPage.CurrentY + Delta * 3)
               Else
                  PrtPage.Line (LeftX, PrtPage.CurrentY + Delta * 3)-(GrRight, PrtPage.CurrentY + Delta * 3)
               End If
               
            End If
         
            GrRow = GrRow + 1
                     
            If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE Then
               Exit Do
            End If
         
         Else
            GrRow = GrRow + 1

         End If
         
      Loop
      
      'imprimimos las líneas verticales
      If Grid.GridLines <> 0 Then
         x = LeftX
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
         
         If Grid.FixedRows = 0 Then
            PrtPage.Line (x, TopSinNombresCampos)-(x, PrtPage.CurrentY + Delta)
         Else
            PrtPage.Line (x, TopY)-(x, PrtPage.CurrentY + Delta)
         End If
         
'         For i = ColIni To ColFin
         For n = 0 To Grid.Cols - 1
            i = ColsPrtFranja(n)
            If i < 0 Or i > ColFin Then
               Exit For
            End If
            
            If lColWi(i) <> 0 Then
               x = x + lColWi(i)
               
'               If x > GrRight Then
'                  Exit For
'               End If
               
               
               If Grid.FixedRows = 0 Then
                  PrtPage.Line (x + Delta, TopSinNombresCampos)-(x + Delta, PrtPage.CurrentY)
               Else
                  PrtPage.Line (x + Delta, TopY + TituloCompartido(i) * 2.5 * (CellHeight * fY))-(x + Delta, PrtPage.CurrentY)
               End If
            End If
            
            
         Next n
         
         RightColX = x + Delta   'posición de la línea vertical más a la derecha
         
      End If
      
      'imprimimos línea horizontal inferior
      If Grid.GridLines <> 0 Then
         PrtPage.Line (LeftX, PrtPage.CurrentY)-(GrRight, PrtPage.CurrentY)
      End If
      
      'If GrRow < Grid.Rows Then
      'si aún no se terminan las líneas o ya se terminaron y falta imprimir más de una linea de totales y no cabe en lo que queda de página, pasamos a página siguiente
      If GrRow < Grid.rows Or (GrRow = Grid.rows And NTotLines > 1 And PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE - PG_LINE_HE * NTotLines) Then
         Dim bHayMas As Boolean, uRow As Integer
         
         uRow = GrRow + 30
         If uRow >= Grid.rows - 1 Then
            uRow = Grid.rows - 1
         End If
         
         bHayMas = False
         For i = GrRow To uRow
            If Grid.TextMatrix(i, ColObligatoria) <> "" Then
               bHayMas = True
               Exit For
            End If
         Next i
      
         If bHayMas = False Then
            Exit Do
         End If
      
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
         Pag = Pag + 1
         
         If PrtPage Is Nothing Then ' no más
            Exit Do
         End If
      
      Else
         Exit Do
            
      End If
   Loop
   
   CurY = PrtPage.CurrentY
   PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY
   
   'imprimimos los totales
   If NTotLines > 0 Then
      
      If TotFntBold = True Then
         PrtPage.FontBold = True
      End If
      For i = 1 To NTotLines
         Call PrtTotales(PrtPage, LeftX, lTotal, lColWi, Grid, i)
         PrtPage.CurrentY = PrtPage.CurrentY + CellHeight * fY * 2
      Next i
      PrtPage.FontBold = False
      
      'imprimimos línea horizontal inferior y líneas verticales de totales
      PrtPage.Line (LeftX, CurY)-(LeftX, PrtPage.CurrentY)
      'PrtPage.Line (GrRight, CurY)-(GrRight, PrtPage.CurrentY)
      'PrtPage.Line (LeftX, PrtPage.CurrentY)-(GrRight, PrtPage.CurrentY)
      PrtPage.Line (RightColX, CurY)-(RightColX, PrtPage.CurrentY)
      PrtPage.Line (LeftX, PrtPage.CurrentY)-(RightColX, PrtPage.CurrentY)
   End If
   
   'imprimimos las observaciones
   If Obs <> "" And CntFranja <= 1 Then
      If PrtPage.CurrentY >= PrtPage.Height - PG_FOOTERSIZE - PG_LINE_HE * 3 Then
         
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
         Pag = Pag + 1
         
         If PrtPage Is Nothing Then ' no más
            Exit Function
         End If
         
'         PrtPage.Print
'         PrtPage.Print
'         PrtPage.Print
'         PrtPage.Print

         'imprimimos encabezado de página
         If PrintHeader = True Then
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, False)

         Else
            Call PrtHeader(PrtPage, LeftX, TLeft, RightX, Pag, True)

         End If

      End If

      PrtPage.FontBold = False
      If PrtPage.FontSize < 9 Then
         PrtPage.FontSize = 9
      End If

      PrtPage.Print
      
      AuxObs = Obs
      k = 0
      i = InStr(AuxObs, vbNewLine)
      Do While i > 0
         
         ReDim Preserve LstObs(k)
         LstObs(k) = Left(AuxObs, i - 1)
         
         AuxObs = Mid(AuxObs, i + Len(vbNewLine))
         i = InStr(AuxObs, vbNewLine)
         k = k + 1
      Loop
      
      If AuxObs <> "" Then   'el último trocito
         ReDim Preserve LstObs(k)
         LstObs(k) = AuxObs
      End If
      
      For k = 0 To UBound(LstObs)
      
         If LstObs(k) = "" Then
            PrtPage.Print
         
         Else
            
            AuxObs = LstObs(k) & " "
            
            Do While AuxObs <> ""
                  
               For i = Len(AuxObs) To 1 Step -1
                  If PrtPage.TextWidth(Left(AuxObs, i)) <= grW Then
                     j = i
                     Do While Mid(AuxObs, j, 1) <> " "
                        j = j - 1
                     Loop
                     PrtPage.CurrentX = LeftX
                     PrtPage.Print Left(AuxObs, j)
                     Exit For
                  End If
               Next i
               
               AuxObs = Mid(AuxObs, j + 1)
            
            Loop
         
         End If
         
      Next k
      
   End If
   
   If CallEndDoc = -1 Then
'      Call PrtFooter(PrtPage, "Total Págs. " & Pag, RightX)
'      PrtPage.EndDoc
      
      If ColFin >= Grid.Cols - 1 Then
         Call PrtFooter(PrtPage, "Total Págs. " & Pag, RightX)
         PrtPage.EndDoc
      Else
         Call PrtFooter(PrtPage, "Continua >>>", RightX)
         Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
      End If
      
   ElseIf CallEndDoc = -2 Then    'llama a EndDoc pero pone Continua>>>>
      Call PrtFooter(PrtPage, "Continua >>>", RightX)
      PrtPage.EndDoc
      
   'ElseIf CallEndDoc = -3 Then     '
   ElseIf CallEndDoc = 0 Then     '
   
      If TieneMasDe1Franja Then
         
         If ColFin < Grid.Cols - 1 Then    'no es la última franja
   
            Call PrtFooter(PrtPage, "Continua >>>", RightX)
            Set PrtPage = NewPage(PrtObj) 'PrtPage.NewPage
            
         End If
         
      End If
      
   End If
   '2860036
     If gMembrete.TxtTitMembrete2 <> "" Then
    Call PrtFooterMembreteRigth(PrtPage, gMembrete.TxtTitMembrete2, gMembrete.TxtTexto2, RightX)
    'Call PrtFooterMembreteLeft(PrtPage, gMembrete.TxtTitMembrete1, gMembrete.TxtTexto1, LeftX)
    End If
     If gMembrete.TxtTitMembrete1 <> "" Then
    'Call PrtFooterMembreteRigth(PrtPage, gMembrete.TxtTitMembrete2, gMembrete.TxtTexto2, RightX)
    Call PrtFooterMembreteLeft(PrtPage, gMembrete.TxtTitMembrete1, gMembrete.TxtTexto1, LeftX)
    
    End If
    
    'fin 2860036
    
   PrtPage.FontName = OldFName
   PrtPage.FontBold = OldFBold
   PrtPage.FontSize = OldFSize
   PrtPage.FontUnderline = OldFUnderline
   
   PrtFlexGrid1Membrete = Pag
   
   On Error GoTo 0
   
End Function
'fin 2860036

'2860036
Public Sub PrtFooterMembreteRigth(PrtPage As Object, ByVal StrTitMembrete As String, ByVal StrTextMembrete As String, ByVal RightX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   PrtPage.CurrentY = PrtPage.Height - PG_FOOTERPOS - 1200
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   PrtPage.CurrentX = RightX - 2200
   PrtPage.Print "------------------------------------------"
    PrtPage.CurrentX = RightX - 2200
   PrtPage.Print StrTitMembrete
    PrtPage.CurrentX = RightX - 2200
   PrtPage.Print StrTextMembrete
   
   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

Public Sub PrtFooterMembreteLeft(PrtPage As Object, ByVal StrTitMembrete As String, ByVal StrTextMembrete As String, ByVal LeftX As Integer)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
         
   PrtPage.CurrentY = PrtPage.Height - PG_FOOTERPOS - 1200
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print "------------------------------------------"
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print StrTitMembrete
   PrtPage.CurrentX = LeftX + 1000
   PrtPage.Print StrTextMembrete
   
   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

'fin 2860036

'2860036
Public Function PrtFlexGridMembrete(PrtObj As Object) As Integer
   Dim Pag As Integer
   Dim Msg As Boolean
   
   Pag = -1

   If ColIni = -1 Then
      ColIni = 0
   End If
   
   If ColFin = -1 Then
      ColFin = Grid.Cols - 1
   End If
   
   If FixedCols = -1 Then
      FixedCols = Grid.FixedCols
   End If

   If CallEndDoc = 0 And Not PermitirMasDe1Franja Then
      Solo1Franja = True
   End If

   TieneMasDe1Franja = False
   Msg = True
   
   Do While ColPrt < Grid.Cols - 1
      
      '2860036
      'Pag = PrtFlexGrid1(PrtObj, Msg)
      Pag = PrtFlexGrid1Membrete(PrtObj, Msg)
      'fin 2860036
      
      If Solo1Franja Or Pag < 0 Then
         Exit Do
      End If
      
      If ColPrt < Grid.Cols - 1 Then
         ColIni = ColPrt + 1
         ColFin = -1
         ColPrt = -1
         CntFranja = CntFranja + 1
      End If
      
      Msg = False
   Loop

   ' Para la siguiente llamada
   ColIni = -1
   ColFin = -1
   ColPrt = -1
   Solo1Franja = False
   FixedCols = -1
   CntFranja = -1

   PrtFlexGridMembrete = Pag
   
End Function
'2860036
