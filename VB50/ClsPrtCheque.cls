VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrtCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FNT_TIMES = "Times New Roman"
Private Const FNT_TITLE = "Arial"
Private Const FNT_COURIER = "Courier New"
Private Const MAX_DEFFNTTIT = 3                   'tamaño arreglo de font de titulos por default

Private Const PC_HCHEQUE = 4400        'altura de un cheque
Private Const PC_BORDEIZQ = 2930       'borde izquierdo cheque
Private Const PC_BORDEIZQPCONT = 3119  'borde izquierdo cheque papel continuo

Private Const PC_ALTOHOJACARTA = 15840  'altura de una hoja tamaño carta en pixels

Private Const TP_HOJACARTA = 1         'tipo papel
Private Const TP_PAPELCONTINUO = 2

Private lNombres() As String           'arreglo de nombres (para membrete)
Public TabNombres As Integer           'tab para segunda parte de cada nombre, después del vbTab
Private lFntNombres() As FontDef_t     'arreglo de fonts para Nombres()

Private lTitulos() As String           'arreglo de titulos
Private lFntTitulos() As FontDef_t     'arreglo de fonts para Titulos()

Private lEncabezados() As String       'arreglo de Encabezados()
Private lTabEncabezados As Integer     'tab para separar encabezados con cada vbTab en el texto
Private lFntEncabezados() As FontDef_t 'arreglo de fonts para Encabezados()

Private BuffEnc As String              'buffer de texto para poner bajo los encabezados y antes de la grilla
Private FntBuffEnc As FontDef_t        'font para BuffEnc()


Public UseCourier As Integer        'indica si se debe usar Courier en vez del font de la grilla

Public gNuevoFormatoCheque As Boolean  'Indica si es el nuevo formato de cheques con celdas para los números

'datos cheque
Public Fecha As Long
Public NominativoA As String
Public Ref As String
Public Banco As String
Public NumCheque As String
Public Valor As Double
Public Lugar As String                 'para anteponerlo a la fecha, en el cheque (ej. Santiago, 24 enero 2009)
Public NumEgreso As String             'para título y lateral cheque
Public PrtDemo As Boolean              'indica si es producto en demo, para que imprima la palabra DEMO en grande

'font default
Private DefFntTitulos(MAX_DEFFNTTIT) As FontDef_t  'almacena la definición de font default de los primeros cuatro títulos
Private DefFntNombre As FontDef_t      'almacena la definición de font default de los nombres (igual para todos)
Private DefFntEncabezado As FontDef_t  'almacena la definición de font default para el encabezado
Private DefFntBuffEnc As FontDef_t     'almacena la definición de font default para BuffEnc

'clase de impresion de grilla si se desea imprimir voucher (opcional)

Public PrtGrid As ClsPrtFlxGrid     'puede ser una MsFlexGrid o una EdFlexGrid (Fairware)

Public TipoPapel As Integer         'TP_HOJACARTA, TP_PAPELCONTINUO

Public AlturaCheque As Long      'altura del cheque desde borde inferior (para hoja carta)
Public BordeIzqCheque As Long    'borde izquierdo del cheque

Public BorrarALaOrden As Boolean    'para hacer cheques nominativos
Public BorrarAlPortador As Boolean

Public BordeIzqChequePCont As Long    'borde izquierdo del cheque para papel continuo
Public BordeSuperiorPCont As Long     'borde superior del cheque para papel continuo

Public BorrarALaOrdenPCont As Boolean    'para hacer cheques nominativos papel continuo
Public BorrarAlPortadorPCont As Boolean

Public BajarValDig As Long            'bajar valor en dígitos para hoja carta
Public MoverValDig As Long            'mover izq/der valor en dígitos para hoja carta
Public BajarFecha As Long             'bajar fecha para hoja carta
Public MoverFecha As Long             'mover izq/der fecha para hoja carta
Public Omitir2DigAno As Boolean          'omitir los 2 primeros dígitos del año (en este caso el 20)
Public MoverAno As Long               'mover el año (2016 o 16) a izq/der

Public BajarOrdenDe As Long           'bajar orden de para hoja carta
Public MoverOrdenDe As Long           'Mover izq/der orden de para hoja carta

Public BajarLaSuma As Long           ' 5 jul 2021: bajar monto en palabras
Public MoverLaSuma As Long           ' 5 jul 2021: Mover monto en palabras

Public BajarValDigPCont As Long       'bajar valor en dígitos para papel continuo
Public MoverValDigPCont As Long       'mover izq/der valor en dígitos para papel continuo
Public BajarFechaPCont As Long        'bajar fecha para papel continuo
Public MoverFechaPCont As Long        'mover izq/der fecha para papel continuo
Public Omitir2DigAnoPCont As Boolean     'omitir los 2 primeros dígitos del año (en este caso el 20) para papel continuo
Public MoverAnoPCont As Long           'mover el año (2016 o 16) a izq/der para papel continuo
Public BajarOrdenDePCont As Long      'bajar orden de para papel continuo
Public MoverOrdenDePCont As Long      'mover izq/der orden de para papel continuo
Public BajarLaSumaPCont As Long           ' 5 jul 2021: bajar monto en palabras
Public MoverLaSumaPCont As Long           ' 5 jul 2021: Mover monto en palabras

Public Function PrintCheque(PrtObj As Object) As Integer
   Dim PrtPage As Object
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   Dim OldFUnderline As Boolean
   Dim OldOrient As Integer
   Dim i As Integer
   Dim TLeft As Integer
   Dim LeftX As Long, TopY As Long
   Dim RightX As Long
   Dim Titulos(0) As String
   Dim Encabezados(4) As String

   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
    
   If TipoPapel = TP_HOJACARTA And PrtPage.Height > PC_ALTOHOJACARTA + 1100 Then
      MsgBox1 "Debe seleccionar tamaño carta en la impresora.", vbExclamation
      Exit Function
   End If

   'margen izquierdo
   TLeft = 12
   
   'margen derecho
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
      
      If Err Then
         MsgBox "Error " & Err & ", " & Error, vbExclamation
      End If
   End If
   
   On Error GoTo 0
   
   Titulos(0) = "EGRESO"
   If NumEgreso <> "" Then
      Titulos(0) = Titulos(0) & " N° " & NumEgreso
   End If
   
   lTitulos = Titulos
   
   Encabezados(0) = "N° Docto.:" & vbTab & NumCheque
   Encabezados(1) = "Fecha:" & vbTab & Format(Fecha, "dd mmm yyyy")
   Encabezados(2) = "Nombre:" & vbTab & NominativoA
   Encabezados(3) = "Valor:" & vbTab & "$ " & Format(Valor, NUMFMT)
   
   If Ref <> "" Then
      Encabezados(4) = "Ref.:" & vbTab & Ref
   End If
      
   lEncabezados = Encabezados
   lTabEncabezados = GetPrtTextWidth("N° Docto.:ww")
   
   'ponemos los fonts por default si corresponde  (debe estar después de la definición de títulos y encabezados por el tamaño de los arreglos)
   Call SetDefaultFonts
   
   If TipoPapel = TP_HOJACARTA Then
      Call PrtHeader(PrtObj, LeftX, TLeft, RightX, False)

      PrtObj.Print
      PrtObj.Print "."
   
      If Not PrtGrid Is Nothing Then
         PrtGrid.PrintHeader = False
         PrtGrid.EsContinuacion = True
         PrtGrid.CallEndDoc = False
         
         Call PrtGrid.PrtFlexGrid(PrtObj)
         
         PrtGrid.PrintHeader = True
         PrtGrid.EsContinuacion = False
         PrtGrid.CallEndDoc = True
      End If
         
      Call PrtFooter(PrtObj)
   
      Call PrtCheque(PrtObj)
      
   Else
   
      Call PrtFooterPapelContinuo(PrtObj)
      Call PrtChequePapelContinuo(PrtObj)
   
   End If
   
   PrtPage.FontName = OldFName
   PrtPage.FontBold = OldFBold
   PrtPage.FontSize = OldFSize
   PrtPage.FontUnderline = OldFUnderline

   PrtObj.EndDoc
   
End Function

Private Sub PrtHeader(PrtPage As Object, LeftX As Long, ByVal TLeft As Long, ByVal RightX As Long, Optional ByVal TinyHeader As Boolean = False)
   Dim i As Integer
   Dim CurX As Integer
   Dim CurY As Integer
   Dim OldFSize As Integer
   Dim OldFName As String
   Dim OldForeColor As Long

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
   
      For i = 0 To UBound(lNombres)
   
         Call SetFont(PrtPage, lFntNombres(i))
         PrtPage.CurrentX = CurX
   
         If lNombres(i) <> "" Then
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
            Call PrtTxtTabs(PrtPage, lEncabezados(i), False, lTabEncabezados, LeftX, RightX)
         End If
      Next i
      
      PrtPage.Print
      
      If BuffEnc <> "" Then
                  
         Call SetFont(PrtPage, FntBuffEnc)
         
         PrtPage.Print
         PrtPage.Print
         
         PrtPage.CurrentX = LeftX
   
         Call PrtBuf(vbLeftJustify, LeftX, 11000, BuffEnc, PrtPage)
      End If
      
   End If
   
End Sub

Public Sub PrtFooter(PrtPage As Object)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
   Dim CurX As Long, CurY As Long
   Dim ArrNomA(10) As String
         
   CurY = PrtPage.Height - AlturaCheque - 1300
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   CurX = 1000
   
   PrtPage.Line (CurX, CurY)-(CurX + 1300, CurY)
   PrtPage.Line (CurX + 4000, CurY)-(CurX + 4000 + 1300, CurY)
   PrtPage.Line (CurX + 8000, CurY)-(CurX + 8000 + 1300, CurY)

   PrtPage.CurrentY = CurY + 60
   PrtPage.CurrentX = CurX + 350
   PrtPage.Print "Preparó";
   PrtPage.CurrentX = CurX + 4000 + 350
   PrtPage.Print "Revisó";
   PrtPage.CurrentX = CurX + 8000 + 400
   PrtPage.Print "Vo Bo";
   
   PrtPage.FontSize = 8
   
   'CurY = CurY + 1700
   CurY = PrtPage.Height - AlturaCheque + 550
   CurX = 100
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY
   PrtPage.Print "N° Egreso: ";
   PrtPage.CurrentX = CurX + 800
   If NumEgreso <> "" Then
      PrtPage.Print NumEgreso
   End If
     
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 200
   PrtPage.Print "Fecha:";
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print Format(Fecha, EDATEFMT)
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 400
   PrtPage.Print "Valor:";
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print "$ " & Format(Valor, NUMFMT)
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 600
   PrtPage.Print "Nombre:";
   PrtPage.CurrentX = CurX + 800
   
   Call SplitStr(NominativoA, 16, ArrNomA())
   PrtPage.Print Left(ArrNomA(0), 16)
   PrtPage.CurrentY = CurY + 800
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print Left(ArrNomA(1), 16)
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 1200
   PrtPage.Print "Retiró Cheque:";
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 1600
   PrtPage.Print "Nombre:";
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print "_____________________"
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 2000
   PrtPage.Print "RUT:";
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print "_____________________"
   
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY + 2400
   PrtPage.Print "Firma:";
   PrtPage.CurrentX = CurX + 800
   PrtPage.Print "_____________________"

   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

Public Sub PrtFooterPapelContinuo(PrtPage As Object)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
   Dim CurX As Long, CurY As Long
   Dim ArrNomA(10) As String
   Dim ArrRef(10) As String
   Dim DeltaY As Integer
         
   CurY = BordeSuperiorPCont
   CurX = 964       '1.7 cm     'inicio de textos de datos
   DeltaY = 170
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
      
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY
     
   PrtPage.CurrentX = CurX - 681                '-1.2 cm
   PrtPage.CurrentY = CurY + 1701 + DeltaY      '3 cm
   'PrtPage.Print "Valor:";
   PrtPage.CurrentX = CurX + 290                '0.5 cm
   PrtPage.Print Format(Valor, NUMFMT)
   
   PrtPage.CurrentX = CurX - 681                '-1.2 cm
   PrtPage.CurrentY = CurY + 2155 + DeltaY      '3.8 cm
   'PrtPage.Print "Fecha:";
   PrtPage.CurrentX = CurX + 290
   PrtPage.Print Format(Fecha, EDATEFMT)
   
   
   PrtPage.CurrentX = CurX - 681                '-1.2 cm
   PrtPage.CurrentY = CurY + 2608 + DeltaY      '4.6 cm
   'PrtPage.Print "Nombre:";
   PrtPage.CurrentX = CurX
   
   Call SplitStr(NominativoA, 16, ArrNomA())
   PrtPage.Print Left(ArrNomA(0), 16)
   PrtPage.CurrentY = CurY + 2838 + DeltaY
   PrtPage.CurrentX = CurX
   PrtPage.Print Left(ArrNomA(1), 16)
   
   PrtPage.CurrentX = CurX - 681                '-1.2 cm
   PrtPage.CurrentY = CurY + 3119 + DeltaY      '5.5 cm
   'PrtPage.Print "Ref.:";
   PrtPage.CurrentX = CurX
   
   Call SplitStr(Ref, 16, ArrRef())
   PrtPage.Print Left(ArrRef(0), 16)
   PrtPage.CurrentY = CurY + 3349 + DeltaY
   PrtPage.CurrentX = CurX
   PrtPage.Print Left(ArrRef(1), 16)
   
   
'   PrtPage.CurrentX = CurX
'   PrtPage.CurrentY = CurY + 1200
'   PrtPage.Print "Retiró Cheque:";
'
'   PrtPage.CurrentX = CurX
'   PrtPage.CurrentY = CurY + 1600
'   PrtPage.Print "Nombre:";
'   PrtPage.CurrentX = CurX + 800
'   PrtPage.Print "_____________________"
'
'   PrtPage.CurrentX = CurX
'   PrtPage.CurrentY = CurY + 2000
'   PrtPage.Print "RUT:";
'   PrtPage.CurrentX = CurX + 800
'   PrtPage.Print "_____________________"
'
'   PrtPage.CurrentX = CurX
'   PrtPage.CurrentY = CurY + 2400
'   PrtPage.Print "Firma:";
'   PrtPage.CurrentX = CurX + 800
'   PrtPage.Print "_____________________"

   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

Public Sub PrtCheque(PrtPage As Object)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
   Dim CurX As Long, CurY As Long
   Dim TxtValor As String, Aux1 As String, Aux2 As String
   Dim i As Long, j As Long
   Dim ChrValor As String
   Dim yFecha As Integer, xFecha As Integer, FmtAno As String
   Dim StrFecha As String
         
   CurY = PrtPage.Height - AlturaCheque
   CurX = BordeIzqCheque
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   'marca de borde superior izquierdo de cheque
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY
   'PrtPage.Print "+"         'marca margen cheque
   
   'valor
   ChrValor = "*"
   PrtPage.CurrentX = CurX + 6520 + MoverValDig
   PrtPage.CurrentY = CurY + 250 + BajarValDig
'   PrtPage.Print Right(String(15, ChrValor) & Format(Valor, NUMFMT) & ChrValor & ChrValor, 17)
   PrtPage.Print String(3, ChrValor) & Format(Valor, NUMFMT) & String(3, ChrValor)
   
   'fecha
   PrtPage.CurrentX = CurX + 4500 + MoverFecha
   xFecha = PrtPage.CurrentX
   PrtPage.CurrentY = CurY + 500 + 600 + BajarFecha
   yFecha = PrtPage.CurrentY
   'PrtPage.Print Left(Lugar, 10) & ", " & Format(Fecha, "dd") & " de " & gNomMes(Month(Fecha)) & " de " & Format(Fecha, "yyyy")
   StrFecha = Left(Lugar, 10) & ", " & Right("0" & Day(Fecha), 2) & " " & Right("0" & Month(Fecha), 2)  ' 26 oct 2016: se pone el mes en numero en vez de texto
   PrtPage.Print StrFecha;
   PrtPage.CurrentX = PrtPage.CurrentX + 100 + MoverAno
   PrtPage.CurrentY = yFecha
   FmtAno = "yyyy"
   If Omitir2DigAno Then
      FmtAno = "yy"
   End If
   PrtPage.Print Format(Fecha, FmtAno)
   
   
   If BorrarALaOrden Then
      'borrar a la orden y al portador
      PrtPage.CurrentX = CurX + 650
      PrtPage.CurrentY = CurY + 950 + 500 + BajarOrdenDe
      PrtPage.Print "xxxxxx"
   End If
   
   If BorrarAlPortador Then
      PrtPage.CurrentX = CurX + 7850
      PrtPage.CurrentY = CurY + 1100 + 450 + BajarOrdenDe
      PrtPage.Print "xxxxxxxx"
   End If
   
   'nombre
   PrtPage.CurrentX = CurX + 1350 + MoverOrdenDe
   PrtPage.CurrentY = CurY + 950 + 500 + BajarOrdenDe
   PrtPage.Print NominativoA
   
   'valor en letras
   PrtPage.CurrentX = CurX + 1500 + MoverLaSuma    ' 5 jul 2021: se agrega MoverLaSuma y BajarLaSuma
   PrtPage.CurrentY = CurY + 1500 + 420 + BajarLaSuma
   TxtValor = Num2Words(Valor, True) & " pesos " & StringStr(150, "- ")
   Aux1 = TxtValor
   Aux2 = ""
   i = Len(TxtValor)
   j = i
   Do While GetPrtTextWidth(Aux1, i = j) > 7500 ' 24 jun 2020: Aux2="" para que si falla muestre sólo un mensaje
      i = i - 1
      Do While Mid(Aux1, i, 1) <> " " And i > 0
         i = i - 1
      Loop
      Aux1 = Left(Aux1, i)
      Aux2 = Mid(TxtValor, i + 1)
   Loop
   
   If Aux2 = "" Then
      PrtPage.Print Aux1
   Else
      PrtPage.Print Aux1
      PrtPage.CurrentX = CurX + 650 + MoverLaSuma
      PrtPage.CurrentY = CurY + 1700 + 500 + BajarLaSuma
      i = Len(Aux2)
      j = i
      Do While GetPrtTextWidth(Aux2, i = j) > 7600 ' 24 jun 2030: para que si falla muestre sólo un mensaje
         i = i - 1
         Aux2 = Left(Aux2, i)
      Loop
      PrtPage.Print Aux2
   End If

   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub

Public Sub PrtChequePapelContinuo(PrtPage As Object)
   Dim TmpFName As String
   Dim TmpFBold As Integer
   Dim TmpFSize As Single
   Dim CurX As Long, CurY As Long
   Dim TxtValor As String, Aux1 As String, Aux2 As String
   Dim i As Long, j As Long
   Dim ChrValor As String, FmtAno As String
   Dim xFecha As Integer, yFecha As Integer
   Dim StrFecha As String
         
   CurY = BordeSuperiorPCont
   CurX = BordeIzqChequePCont
   
   TmpFName = PrtPage.FontName
   TmpFBold = PrtPage.FontBold
   TmpFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   'marca de borde superior izquierdo de cheque
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY
   'PrtPage.Print "+"         'marca margen cheque
   
   'valor
   ChrValor = "*"
   
   PrtPage.CurrentX = CurX + 6804 + MoverValDigPCont    '12
   PrtPage.CurrentY = CurY + 794 + BajarValDigPCont      '1.4
   PrtPage.Print Right(String(15, ChrValor) & Format(Valor, NUMFMT) & ChrValor & ChrValor, 17)
   
   'fecha
   PrtPage.CurrentX = CurX + 3402 + MoverFechaPCont '6
   xFecha = PrtPage.CurrentX
   PrtPage.CurrentY = CurY + 1418 + BajarFechaPCont '2.5
   yFecha = PrtPage.CurrentY
'   PrtPage.Print Left(Lugar, 10) & ", " & Format(Fecha, "dd");
   StrFecha = Left(Lugar, 10) & ", " & Right("0" & Day(Fecha), 2) & " " & Right("0" & Month(Fecha), 2)
   PrtPage.Print StrFecha;
   PrtPage.CurrentX = PrtPage.CurrentX + 100 + MoverAno
   PrtPage.CurrentY = yFecha
   FmtAno = "yyyy"
   If Omitir2DigAno Then
      FmtAno = "yy"
   End If
   PrtPage.Print Format(Fecha, FmtAno)
   
'   PrtPage.CurrentX = CurX + 5953   '10.5
''   PrtPage.Print gNomMes(Month(Fecha));
'   PrtPage.Print Right("0" & Month(Fecha), 2);
'
'   PrtPage.CurrentX = CurX + 7938 + MoverAnoPCont  '14
'   FmtAno = "yyyy"
'   If Omitir2DigAnoPCont Then
'      FmtAno = "yy"
'   End If
'   PrtPage.Print Format(Fecha, FmtAno)
   
   
'   If BorrarALaOrden Then
'      'borrar a la orden y al portador
'      PrtPage.CurrentX = CurX + 650
'      PrtPage.CurrentY = CurY + 950 + 500
'      PrtPage.Print "xxxxxx"
'   End If
'
'   If BorrarAlPortador Then
'      PrtPage.CurrentX = CurX + 7850
'      PrtPage.CurrentY = CurY + 1100 + 450
'      PrtPage.Print "xxxxxxxx"
'   End If
   
   'nombre
   PrtPage.CurrentX = CurX + 1417 + MoverOrdenDePCont  '2.5  ' 5 jul 2021: se agrega MoverLaSuma y BajarLaSuma
   PrtPage.CurrentY = CurY + 1984 + BajarOrdenDe '3.5
   PrtPage.Print NominativoA
   
   'valor en letras
   PrtPage.CurrentX = CurX + 1417 + MoverLaSumaPCont    '2.5
   PrtPage.CurrentY = CurY + 2438 + BajarLaSumaPCont    '4.3
   TxtValor = Num2Words(Valor, True) & " pesos " & StringStr(150, "- ")
   Aux1 = TxtValor
   Aux2 = ""
   i = Len(TxtValor)
   j = i
   Do While GetPrtTextWidth(Aux1, i = j) > 7370
      i = i - 1
      Do While Mid(Aux1, i, 1) <> " " And i > 0
         i = i - 1
      Loop
      Aux1 = Left(Aux1, i)
      Aux2 = Mid(TxtValor, i + 1)
   Loop
   
   If Aux2 = "" Then
      PrtPage.Print Aux1
   Else
      PrtPage.Print Aux1
      PrtPage.CurrentX = CurX + 567 + MoverLaSumaPCont    '1.0
      PrtPage.CurrentY = CurY + 2665 + BajarLaSumaPCont  '4.7
      i = Len(Aux2)
      j = i
      Do While GetPrtTextWidth(Aux2, i = j) > 7370
         i = i - 1
         Aux2 = Left(Aux2, i)
      Loop
      PrtPage.Print Aux2
   End If

   PrtPage.FontName = TmpFName
   PrtPage.FontBold = TmpFBold
   PrtPage.FontSize = TmpFSize

End Sub
Public Sub Clear()
               
   UseCourier = False
   
   If AlturaCheque = 0 Then
      AlturaCheque = PC_HCHEQUE
   End If
   
   If BordeIzqCheque = 0 Then
      BordeIzqCheque = PC_BORDEIZQ
   End If
      
   'BordeSuperiorPCont=0
      
   If BordeIzqChequePCont = 0 Then
      BordeIzqChequePCont = PC_BORDEIZQPCONT
   End If

   
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
   
   PrtDemo = False
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
Private Sub Class_Initialize()
   
   ReDim lNombres(0)
   ReDim lFntNombres(0)
   ReDim lTitulos(0)
   ReDim lFntTitulos(0)
   ReDim lEncabezados(0)
   ReDim lFntEncabezados(0)
   
   TipoPapel = TP_HOJACARTA
      
   Call Clear

End Sub

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

   If FntBuffEnc.FontName = "" Then
      FntBuffEnc.FontName = DefFntBuffEnc.FontName
      FntBuffEnc.FontSize = DefFntBuffEnc.FontSize
      FntBuffEnc.FontBold = DefFntBuffEnc.FontBold
      FntBuffEnc.FontUnderline = DefFntBuffEnc.FontUnderline
   End If
   
End Function

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
         PrtPage.Print Txt
         'Call PrtBuf(vbLeftJustify, CurX, RightX, Txt, PrtPage)
         'PrtPage.Print
      End If
      
   End If

End Sub

Public Sub PrtMarca(PrtObj As Object)
   Dim PrtPage As Object
   Dim CurX As Long, CurY As Long
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   
   Set PrtPage = Nothing
   Set PrtPage = GetPrtPage(PrtObj)
         
   If PrtPage.Height > PC_ALTOHOJACARTA + 1100 Then
      MsgBox1 "Debe seleccionar tamaño carta en la impresora.", vbExclamation
      Exit Sub
   End If
         
   CurY = PrtPage.Height - AlturaCheque
   CurX = BordeIzqCheque
   
   OldFName = PrtPage.FontName
   OldFBold = PrtPage.FontBold
   OldFSize = PrtPage.FontSize
   
   PrtPage.FontName = FNT_TITLE
   PrtPage.FontBold = False
   PrtPage.FontSize = 10
   
   'marca de borde superior izquierdo de cheque
   PrtPage.CurrentX = CurX
   PrtPage.CurrentY = CurY
   PrtPage.Print "+"         'marca margen cheque
   
   PrtPage.FontName = OldFName
   PrtPage.FontBold = OldFBold
   PrtPage.FontSize = OldFSize
   
   PrtObj.EndDoc
   
End Sub

