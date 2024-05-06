Attribute VB_Name = "PamFGrid"
Option Explicit

' Por defecto se asume que:
' ID está en Cols-2        - penúltima columna
' Estado está en Cols-1    - última columna
'

Public Const FGR_MAXROWS = 32760 ' en realidad son 32768

' Anchos para Grillas
Public Const FW_FECHA = 1050
Public Const FW_RUT = 1050

' Acciones y Estados
Public Const FGR_I = "I"   ' Insert
Public Const FGR_U = "U"   ' Update
Public Const FGR_D = "D"   ' Delete

Public GrRow As Integer
Public GrCol As Integer

' bIncludeCero: indica que debe incluir las columnas de ancho 0 y las filas de alto 0
Function FGr2Html(Grid As Control, ByVal fname As String, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False) As Boolean
   Dim oRow As Integer, oCol As Integer, Gr As Control
   Dim r As Integer, c As Integer
   Dim Mouse As Integer
   Dim Buf As String, BAlign As String
   Dim Fd As Long

   On Error Resume Next
   
   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If
   
   FGr2Html = False
   
   Mouse = Gr.MousePointer
   Gr.MousePointer = vbHourglass
   DoEvents
   
   oRow = Gr.Row
   oCol = Gr.Col
   
   Title = Trim(Title)

   Fd = FreeFile()
   Open fname For Output As #Fd
   If ERR Then
      MsgErr fname
      Exit Function
   End If
   
   Buf = "<html><head><title>" & Title & "</title></head><body><basefont face=Arial>"
   Print #Fd, Buf
   
   Buf = "<p align=center><b>" & Title & "</b></p><table align=center border=1 cellSpacing=1>"
   Print #Fd, Buf
   
   BAlign = "<ColGroup>"
   For c = 0 To Gr.Cols - 1
      If Gr.ColWidth(c) > 0 Then
         If Gr.ColAlignment(c) = flexAlignRightCenter Then
            BAlign = BAlign & "<col Align=right>"
         ElseIf Gr.ColAlignment(c) = flexAlignCenterCenter Then
            BAlign = BAlign & "<col Align=center>"
         Else
            BAlign = BAlign & "<col Align=left>"
         End If
      End If
   Next c
   BAlign = BAlign & "</ColGroup>"
         
   For r = 0 To Gr.rows - 1
      
      If r = Gr.FixedRows Then
         Print #Fd, BAlign
      End If
      
      If Gr.RowHeight(r) > 0 Then
      
         Buf = "<tr>"
         For c = 0 To Gr.Cols - 1
            If Gr.ColWidth(c) > 0 Then
               If Trim(Gr.TextMatrix(r, c)) <> "" Then
                  If r < Gr.FixedRows Or c < Gr.FixedCols Then
                     Buf = Buf & "<th>" & Gr.TextMatrix(r, c) & "</th>"
                  Else
                     Buf = Buf & "<td>" & Gr.TextMatrix(r, c) & "</td>"
                  End If
               Else
                  Buf = Buf & "<td>&nbsp;</td>"
               End If
            End If
         Next c
         Buf = Buf & "</tr>"
         
         Print #Fd, Buf
      End If
   Next r
   
   Buf = "</table><p align=right><font size=1>Generado por " & App.Title & "<br>" & FmtFecha(Now) & Format(Now, " h:nn") & "</font></p></basefont></body></html>"
   Print #Fd, Buf
   
   Close #Fd
   
   Gr.Row = oRow
   Gr.Col = oCol

   Gr.MousePointer = Mouse

End Function

' bIncludeCero: indica que debe incluir las columnas de ancho 0 y las filas de alto 0
' Si se indica ColOblig, sólo incluye las filas que tienen datos en esa columna.
Public Function FGr2String(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False, Optional ByVal ColOblig As Integer = -1, Optional ByVal bTraspose As Boolean = False, Optional ByVal bNoQuotes As Boolean) As String
   Dim fxr As Integer, Gr As Control
   Dim FxC As Integer
   Dim r As Integer, c As Integer
   Dim Mouse As Long
   Dim Clip As String

   On Error Resume Next

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   Mouse = Gr.MousePointer
   Gr.MousePointer = vbHourglass
   DoEvents
   
   r = Gr.Row
   c = Gr.Col
   
   fxr = Gr.FixedRows
   FxC = Gr.FixedCols
   
   Gr.FixedRows = 0
   Gr.FixedCols = 0

   Title = Trim(Title)
   If Title <> "" Then
      Title = Title & vbCrLf ' 10 jun 2019: cambia de vCr a vCrLf
   End If

   If bIncludeCero And ColOblig = -1 Then
      Gr.Row = 0
      Gr.Col = 0

      Gr.RowSel = Gr.rows - 1
      Gr.ColSel = Gr.Cols - 1
      
      If bTraspose = False Then
         Clip = Gr.Clip
      Else
         Debug.Print "FGr2String: *** NO IMPLEMENTADO ***"
      End If
   ElseIf bTraspose Then
      Clip = FGrClipNoCeroTrasp(Gr, ColOblig, fxr)
   Else
      Clip = FGrClipNoCero(Gr, ColOblig, fxr)
   End If

   If bNoQuotes Then
      Clip = Replace(Clip, """", "")
   End If

   Gr.FixedRows = fxr
   Gr.FixedCols = FxC
   
   Gr.Row = r
   Gr.Col = c

   FGr2String = Title & Clip
   
   Gr.MousePointer = Mouse

End Function

Public Function FGr2Clip(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False, Optional ByVal bNoQuotes As Boolean) As Long
   Dim Clip As String

   Clip = FGr2String(Grid, Title, bIncludeCero, , False, bNoQuotes)

   FGr2Clip = SetClipText(Clip)

End Function
Public Function FGr2ClipTrasp(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False, Optional ByVal bNoQuotes As Boolean) As Long
   Dim Clip As String

   Clip = FGr2String(Grid, Title, bIncludeCero, , True, bNoQuotes)

   FGr2ClipTrasp = SetClipText(Clip)

End Function

Public Function FGrClipNoCeroTrasp(Gr As Control, Optional ByVal RowOblig As Integer = -1, Optional ByVal FixedCols As Integer = 0) As String
   Dim Buf As String, lBuf As String
   Dim c As Integer, r As Integer, bInc As Boolean

   Buf = ""
   For r = 0 To Gr.Cols - 1
      bInc = True
   
      If Gr.ColWidth(r) <= 0 Then
         bInc = False
      ElseIf RowOblig >= 0 Then
         If Gr.TextMatrix(RowOblig, r) = "" And r >= FixedCols Then
            bInc = False
         End If
      End If
      
      
      If bInc Then
         lBuf = ""
         
         If Gr.MergeCol(r) Then
            For c = 0 To Gr.rows - 1
               If Gr.RowHeight(c) > 0 Then
                  lBuf = lBuf & Gr.TextMatrix(c, r) & vbTab
'                  If Len(Gr.TextMatrix(r, c)) > 0 Then
'                     Exit For  ' sólo pone el primero porque las otras son iguales
'                  End If
               End If
            Next c
         Else
            For c = 0 To Gr.rows - 1
               If Gr.RowHeight(c) > 0 Then
                  lBuf = lBuf & Gr.TextMatrix(c, r) & vbTab
               End If
            Next c
         End If
         Buf = Buf & lBuf & vbCr
      End If
   Next r

   FGrClipNoCeroTrasp = Buf
   
End Function
' No incluye columnas de ancho cero
Public Function FGrClipNoCero(Gr As Control, Optional ByVal ColOblig As Integer = -1, Optional ByVal FixedRows As Integer = 0) As String
   Dim Buf As String, LinBuf As String, bLin As Boolean, LastLin As Long
   Dim c As Integer, r As Integer, bInc As Boolean

   Buf = ""
   For r = 0 To Gr.rows - 1
      bInc = True
   
      If Gr.RowHeight(r) <= 0 Then
         bInc = False
      ElseIf ColOblig >= 0 Then
         If Gr.TextMatrix(r, ColOblig) = "" And r >= FixedRows Then
            bInc = False
         End If
      End If
      
      
      If bInc Then
         LinBuf = ""
         bLin = False  ' 15 jun 2021: para no copiar líneas en blanco
         
         If Gr.MergeRow(r) Then
            For c = 0 To Gr.Cols - 1
               If Gr.ColWidth(c) > 0 Then
                  '2971346
                  'LinBuf = LinBuf & Gr.TextMatrix(r, c) & vbTab
                  LinBuf = LinBuf & RemoveNoPrtChars(Gr.TextMatrix(r, c)) & vbTab
                   '2971346
                  bLin = bLin Or (Gr.TextMatrix(r, c) <> "")
'                  If Len(Gr.TextMatrix(r, c)) > 0 Then
'                     Exit For  ' sólo pone el primero porque las otras son iguales
'                  End If
               End If
            Next c
         Else
            For c = 0 To Gr.Cols - 1
               If Gr.ColWidth(c) > 0 Then
                  '2971346
                  'LinBuf = LinBuf & Gr.TextMatrix(r, c) & vbTab
                  LinBuf = LinBuf & RemoveNoPrtChars(Gr.TextMatrix(r, c)) & vbTab
               '2971346
                  bLin = bLin Or (Gr.TextMatrix(r, c) <> "")
               End If
            Next c
         End If
         
         Buf = Buf & Replace(LinBuf, Chr(34), "") & vbCrLf ' 10 jun 2019: cambia de vCr a vCrLf
         
         If bLin Then
            LastLin = Len(Buf)
         End If
      End If
               
   Next r

   FGrClipNoCero = Mid(Buf, 1, LastLin)
   
End Function

' Retorna el Texto asociado al item dado, de la columna especificada
' OJO: Sólo para FlexGrid
Public Function FGrCbText(FGrid As Control, ByVal Col As Integer, ByVal Item As Long) As String

   FGrCbText = cbItemText(FGrid.CbList(Col), Item)
   'FGrCbText = FGrid.CbList(Col).List(FindItem(FGrid.CbList(Col), Item))
   
End Function
' Selecciona una fila de la grilla
Public Sub FGrSelRow(Grid As Control, ByVal Row As Integer, Optional ByVal TopRow As Boolean = 1)
   Dim Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   If Row >= Gr.rows Then
      Row = Gr.FixedRows
   End If

   Gr.Row = Row
   Gr.Col = 0
      
   Gr.RowSel = Row
   Gr.ColSel = Gr.Cols - 1
      
   If TopRow = True And Row >= Gr.FixedRows Then

      If Row < Gr.TopRow Then
         Gr.TopRow = Row
      ElseIf Row >= Gr.TopRow + FGrVRows(Gr) - Gr.FixedRows * 3 Then
         Gr.TopRow = Row
      End If
      
   End If

End Sub

Public Function FGrSelCell(Grid As Control, Row As Integer, Col As Integer)
   Dim Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   Gr.Row = Row
   Gr.RowSel = Row
   
   Gr.Col = Col
   Gr.ColSel = Col
      
   If Row >= Gr.FixedRows Then

      If Row < Gr.TopRow Then
         Gr.TopRow = Row
      ElseIf Row >= Gr.TopRow + FGrVRows(Gr) - Gr.FixedRows * 3 Then
         Gr.TopRow = Row
      End If
      
   End If
   
   If Col >= Gr.FixedCols Then

      If Col < Gr.LeftCol Then
         Gr.LeftCol = Col
      End If
      
   End If
   
   If Gr.visible Then
      Gr.SetFocus
   End If
   
End Function

Public Sub FGrInitTitle_(FGr As Control)
   Dim c As Integer

   For c = 0 To FGr.Cols - 1
      FGr.FixedAlignment(c) = flexAlignCenterCenter
   Next c

End Sub

Public Sub FGrSetup(Grid As Control, Optional ByVal BLeft As Boolean = 0)
   Dim i As Integer, Gr As Control
   
   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If
   
   For i = 0 To Gr.Cols - 1
      Gr.FixedAlignment(i) = flexAlignCenterCenter
      
      If BLeft <> 0 And Gr.ColAlignment(i) = flexAlignGeneral Then
         Gr.ColAlignment(i) = flexAlignLeftCenter
      End If
   Next i

   Gr.BackColorBkg = vbWindowBackground

End Sub
' Permite manejar las modificaciones en la grilla
' Acciones posibles: FGR_U, FGR_D
' idRow: Fila que contiene el id (por defecto la penúltima)
' stRow: Fila que contiene el estado (por defecto la última)
' bDelCol: Indica si la función elimina la columna o no
' Para grabar hay que revisar la fila StRow: FGR_I, FGR_U, FGR_D
' OJO: Sólo para FlexGrid
Sub FGrModCol(FGr As Control, ByVal Col As Integer, ByVal Action As String, Optional ByVal IdRow As Integer = -1, Optional ByVal StRow As Integer = -1, Optional ByVal bDelCol As Boolean = 1)

   If IdRow = -1 Then
      IdRow = FGr.rows - 2
   End If
   
   If StRow = -1 Then
      StRow = FGr.rows - 1
   End If

   If FGr.TextMatrix(IdRow, Col) = "" Then ' Nuevo
      If Action = FGR_U Or Action = FGR_I Then
         FGr.TextMatrix(StRow, Col) = FGR_I
      ElseIf Action = FGR_D Then
         If bDelCol Then
            
            FGr.ColWidth(Col) = 0
            FGr.Cols = FGr.Cols + 1
         End If
         
         FGr.TextMatrix(StRow, Col) = ""
      End If
   Else
      If Action = FGR_U And FGr.TextMatrix(StRow, Col) <> FGR_I Then
         FGr.TextMatrix(StRow, Col) = FGR_U
      ElseIf Action = FGR_D Then
         FGr.TextMatrix(StRow, Col) = FGR_D
         If bDelCol Then
            FGr.ColWidth(Col) = 0
            FGr.Cols = FGr.Cols + 1
            
         End If
      End If
   End If

End Sub
' Permite manejar las modificaciones en la grilla
' Acciones posibles: FGR_U, FGR_D
' idCol: Columna que contiene el id (por defecto la penúltima)
' stCol: Columna que contiene el estado (por defecto la última)
' bDelRow: Indica si la función elimina la fila o no
' Para grabar hay que revisar la columna StCol: FGR_I, FGR_U, FGR_D
' OJO: Sólo para FlexGrid
Sub FGrModRow(FGr As Control, ByVal Row As Integer, ByVal Action As String, Optional ByVal idCol As Integer = -1, Optional ByVal StCol As Integer = -1, Optional ByVal bDelRow As Boolean = 1)

   If idCol = -1 Then
      idCol = FGr.Cols - 2
   End If
   
   If StCol = -1 Then
      StCol = FGr.Cols - 1
   End If

   If FGr.TextMatrix(Row, idCol) = "" Then ' Nuevo
      If Action = FGR_U Or Action = FGR_I Then
         FGr.TextMatrix(Row, StCol) = FGR_I
      ElseIf Action = FGR_D Then
         If bDelRow Then
            FGr.RemoveItem Row
            
'            If bDelRow = 1 Then ' matiene el numero de filas
               FGr.rows = FGr.rows + 1
'            End If
         Else
            FGr.TextMatrix(Row, StCol) = ""
         End If
      End If
   Else
      If Action = FGR_U And FGr.TextMatrix(Row, StCol) <> FGR_I Then
         FGr.TextMatrix(Row, StCol) = FGR_U
      ElseIf Action = FGR_D Then
         FGr.TextMatrix(Row, StCol) = FGR_D
         If bDelRow Then
            FGr.RowHeight(Row) = 0
            
'            If DelRow = 1 Then ' matiene el numero de filas visibles
'               FGr.rows = FGr.rows + 1 ' Se comenta para que permita eliminar mas registros en la grilla ADO 2699586 FPG 26-01-2022
'            End If
            
         End If
      End If
   End If

End Sub

Public Sub FGrClear(Gr As Control, Optional ByVal bAll As Boolean = 0)
   Dim r As Integer, fr As Integer
   
   fr = Gr.FixedRows
   r = Gr.rows
   
   If bAll Then
      Gr.rows = 0
   Else
      Gr.rows = fr
   End If
   
   Gr.rows = r
   
   If bAll Then
      Gr.FixedRows = fr
   End If

End Sub
Public Sub FGrClearRow(Gr As Control, ByVal Row As Integer)
   Dim c As Integer
   
   For c = 0 To Gr.Cols - 1
      Gr.TextMatrix(Row, c) = ""
   Next c
   
End Sub
' Busca la primera fila libre o agrega una
Public Function FGrAddRow(Gr As Control, Optional ByVal Col As Integer = 0, Optional ByVal bHeightZero As Boolean = 0)
   Dim r As Integer

   For r = Gr.FixedRows To Gr.rows - 1
      If Gr.TextMatrix(r, Col) = "" Then
         If bHeightZero Or Gr.RowHeight(r) > 0 Then
            FGrAddRow = r
            Exit Function
         End If
      End If
   Next r

   ' Si no habia ninguna libre...
   Gr.rows = Gr.rows + 1
   FGrAddRow = Gr.rows - 1
   
End Function

' Determina y completa el número de filas visibles
Public Function FGrVRows(Gr As Control, Optional ByVal AdicRows As Byte = 0) As Integer
   Dim rows As Integer, hRows As Integer, i As Integer, Hei As Long
   
   Gr.rows = Gr.rows + AdicRows
   
   If Gr.rows > 1 Then
      Hei = Screen.TwipsPerPixelY
      For i = 0 To Gr.rows - 1
         Hei = Hei + Gr.RowHeight(i)
         If Gr.RowHeight(i) Then
            Hei = Hei + Screen.TwipsPerPixelY
            
            If hRows <= 0 Then
               hRows = Gr.RowHeight(i)
            End If

         End If
         
         If Hei > Gr.Height Then
            FGrVRows = Gr.rows
            Exit Function
         End If
         
      Next i
   
   
      rows = Gr.rows + Int((Gr.Height - Hei) / hRows + 1) + 1
   
'      If Gr.RowHeight(1) Then
'         hRows = Gr.RowHeight(1)
'      ElseIf Gr.RowHeight(0) Then
'         hRows = Gr.RowHeight(0)
'      Else
'         hRows = 240
'      End If
   Else
      If Gr.RowHeight(0) Then
         hRows = Gr.RowHeight(0)
      Else
         hRows = 240
      End If
   
      rows = (Gr.Height / hRows)
   End If

   If Gr.rows < rows Then
      Gr.rows = rows + 1
   End If

   FGrVRows = rows

End Function
' Ubica un form justo abajo de la posición de la celda indicada
Public Sub FGrLocateForm(Grid As Control, Frm As Form, ByVal Row As Integer, ByVal Col As Integer)
   Dim Lft As Integer, GTop As Integer, Gr As Control
   Dim Cnt1 As Object

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   ' Primero sumamos los Left de la grilla
   Set Cnt1 = Gr
   Lft = 0
   GTop = W.YCaption + W.yFrame
   On Error Resume Next
   Do
      Lft = Lft + Cnt1.Left
      GTop = GTop + Cnt1.Top
      If Cnt1.Parent Is Nothing Then
         Exit Do
      End If
      
      Set Cnt1 = Cnt1.Container
   
   Loop

   ' On Error Resume Next
   Frm.Left = Lft + Gr.ColPos(Col) + Screen.TwipsPerPixelX
   Frm.Top = GTop + Gr.RowPos(Row) + Screen.TwipsPerPixelY

End Sub
' Ubica un control en la posición de la columna indicada
Public Sub FGrLocateCntrl(Grid As Control, Cnt As Control, ByVal Col As Integer, Optional BRight As Boolean = 0)
   Dim Lft As Integer, Gr As Control
   Dim Cnt1 As Object

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   ' Primero sumamos los Left de la grilla
   Set Cnt1 = Grid
   Lft = 0
   Do
      Lft = Lft + Cnt1.Left
      If Grid.Parent.hWnd = Cnt1.Container.hWnd Then
         Exit Do
      End If
      
      Set Cnt1 = Cnt1.Container
   
   Loop

   ' Restamos los Left del control
   Set Cnt1 = Cnt.Container
   Do Until Grid.Parent.hWnd = Cnt1.hWnd
      Lft = Lft - Cnt1.Left
      
      Set Cnt1 = Cnt1.Container
   
   Loop

   If BRight Then
      Cnt.Left = Lft + Gr.ColPos(Col) + Gr.ColWidth(Col) - Cnt.Width - Screen.TwipsPerPixelX
   Else
      Cnt.Left = Lft + Gr.ColPos(Col) + Screen.TwipsPerPixelX
      Cnt.Width = Gr.ColWidth(Col) + Screen.TwipsPerPixelX
   End If
   
End Sub


'
' Activa la celda sobre la que está el cursor
'
Sub FGrSetCursorCell(Grid As Control, Optional ByVal x As Integer = 0, Optional ByVal Y As Integer = 0)
   Dim Pt As POINTAPI_T, Gr As Control
   Dim RECT As RECT_T
   Dim i As Integer, s As Integer, a As Integer
   
   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If
   
   Gr.Col = Gr.MouseCol
   Gr.Row = Gr.MouseRow
   
   If Gr.visible Then
      Gr.SetFocus
   End If
   
   Exit Sub
   
   If x >= 0 Then
      Pt.x = x
      Pt.Y = Y
   Else
      Call GetCursorPos(Pt)
   
      Call GetWindowRect(Gr.hWnd, RECT)
   
      Pt.x = (Pt.x - RECT.Left) * Screen.TwipsPerPixelX
      Pt.Y = (Pt.Y - RECT.Top) * Screen.TwipsPerPixelY
   End If
   
   If Gr.TopRow = Gr.FixedRows Then
      s = 0
      a = 0
   Else
      a = Gr.TopRow
      s = s + Gr.RowHeight(0) + Screen.TwipsPerPixelY
   End If
   
   For i = a To Gr.rows - 1
      s = s + Gr.RowHeight(i) + Screen.TwipsPerPixelY
      If i >= Gr.FixedRows And Pt.Y < s Then
         Gr.Row = i
         Exit For
      End If
   Next i
   
   If Gr.LeftCol = Gr.FixedCols Then
      s = 0
      a = 0
   Else
      a = Gr.LeftCol
      s = s + Gr.ColWidth(0) + Screen.TwipsPerPixelX
   End If

   For i = a To Gr.Cols - 1
      s = s + Gr.ColWidth(i) + Screen.TwipsPerPixelX
      If i >= Gr.FixedCols And Pt.x < s Then
         Gr.Col = i
         
         Exit For
      End If
   Next i
      
   Gr.ColSel = Gr.Col
   Gr.RowSel = Gr.Row
   
   If Gr.visible Then
      Gr.SetFocus
   End If
      
End Sub
' Permite buscar un registro mientras el usuario va tipeando
Public Sub FGrSeekRow2(Gr As Control, ByVal KeyAscii As Integer, ByVal Col As Integer)
   Static Buf As String, Tm As Double, iRow As Integer

   If KeyAscii = vbKeyEscape Then
      Buf = ""
      Tm = 0
      iRow = 0
   Else
   
      If Now - Tm > TimeSerial(0, 0, 1) Then  ' espera 3 segundos
         Buf = ""
'      Else
'         iRow = -1
      End If
            
      Buf = Buf & Chr(KeyAscii)
      If FGrSeekRow(Gr, Buf, Col, iRow) Then
         iRow = Gr.Row + 1
         Tm = Now
      End If
   End If

End Sub
' Permite buscar un registro de acuerdo a como comienza, ver FGrSeekRow2
Public Function FGrSeekRow(Gr As Control, ByVal Buf As String, ByVal Col As Integer, Optional ByVal iRow As Integer = -1) As Boolean
   Dim r As Integer, l As Integer
   
   FGrSeekRow = False
   
   If Buf = "" Then
      Exit Function
   End If
   
   If iRow = -1 Then
      iRow = Gr.Row
   End If
   
   If iRow < Gr.FixedRows Then
      iRow = Gr.FixedRows
   End If
      
   l = Len(Buf)
   Buf = Trim(Buf)
   
   For r = iRow To Gr.rows - 1
      If StrComp(Left(Gr.TextMatrix(r, Col), l), Buf, vbTextCompare) = 0 Then
         Gr.TopRow = r
         Gr.Row = r
         Call FGrSelCell(Gr, r, Col)
         FGrSeekRow = True
         Exit Function
      End If
   Next r

   ' damos la vuelta por abajo
   If iRow > Gr.FixedRows Then
      For r = Gr.FixedRows To iRow
         If StrComp(Left(Gr.TextMatrix(r, Col), l), Buf, vbTextCompare) = 0 Then
            Gr.TopRow = r
            Gr.Row = r
            Call FGrSelCell(Gr, r, Col)
            FGrSeekRow = True
            Exit Function
         End If
      Next r
   End If

End Function
Function FGrGet(Gr As Control, ByVal Row As Integer, ByVal Col As Integer) As String

   FGrGet = Gr.TextMatrix(Row, Col)

End Function

Public Sub ChGridFontSize(Grid As MSFlexGrid, ByVal SmallFont As Integer)
   Dim i As Integer
   
   If SmallFont <> 0 Then
      Grid.Font.Name = "Arial"
      Grid.Font.Size = 7
            
      For i = 0 To Grid.Cols - 1
         Grid.ColWidth(i) = Grid.ColWidth(i) * 0.8
      Next i
   Else
      Grid.Font.Name = "MS Sans Serif"
      Grid.Font.Size = 8
           
      For i = 0 To Grid.Cols - 1
         If Grid.ColWidth(i) > 0 Then
            Grid.ColWidth(i) = Grid.ColWidth(i) * 1.25
         End If
      Next i
      
   End If

End Sub
'Destaca una columna en una grilla, dándole un font o color distinto del resto
'Parámetros:
'  StyleType:
'     "B": bold
'     "I": italic
'     "FC": color del texto
'     "BC": color de fondo
'     "T": efecto especial de volumen para el texto
'  StyleColor: color del texto o del fondo, para ser usado con StyleType = "BC" o "FC" (opcional)
'  StartCol: columna a partir de la cual se realiza el cambio (opcional)
Public Sub FGrSetColStyle(Grid As Control, ByVal Col As Integer, ByVal StyleType As String, Optional ByVal StyleColor As Long = 0, Optional ByVal StartRow As Integer = 0)
   Dim CurrFillStyle As Integer
   Dim oRow As Integer, oCol As Integer, St As String, i As Integer
   
   CurrFillStyle = Grid.FillStyle
   
   Grid.FillStyle = flexFillRepeat
   
   oRow = Grid.Row
   oCol = Grid.Col
   
   Grid.Col = Col
   Grid.ColSel = Col
   
   Grid.Row = StartRow

   Grid.RowSel = Grid.rows - 1
   
   StyleType = UCase(StyleType)
   
   Do
      i = InStr(StyleType, "+")
      If i > 0 Then
         St = Left(StyleType, i - 1)
         StyleType = Mid(StyleType, i + 1)
      Else
         St = StyleType
      End If
         
      Select Case St
         Case "B":
            Grid.CellFontBold = True
         
         Case "I":
            Grid.CellFontItalic = True
            
         Case "U": ' Underline
            Grid.CellFontUnderline = True
            
         Case "FC":
            Grid.CellForeColor = StyleColor
            
         Case "BC":
            Grid.CellBackColor = StyleColor   'vb3DLight
            
         Case "T":
            Grid.CellTextStyle = flexTextRaisedLight
      End Select
      
   Loop While i > 0
   
   Grid.FillStyle = CurrFillStyle
   
   Grid.Col = oCol
   Grid.ColSel = oCol
   
   Grid.Row = oRow
   Grid.RowSel = oRow

End Sub
'Destaca una línea en una grilla, dándole un font o color distinto del resto
'Parámetros:
'  StyleType:  (Se puede poner B+I+FC)
'     "B": bold
'     "I": italic
'     "FC": color del texto
'     "BC": color de fondo
'     "T": efecto especial de volumen para el texto
'  StyleColor: color del texto o del fondo, para ser usado con StyleType = "BC" o "FC" (opcional)
'  StartCol: columna a partir de la cual se realiza el cambio (opcional)
Public Sub FGrSetRowStyle(Grid As Control, ByVal Row As Integer, ByVal StyleType As String, Optional ByVal Style As Long = 0, Optional ByVal StartCol As Integer = 0, Optional ByVal EndCol As Integer = -1)
   Dim CurrFillStyle As Integer
   Dim oRow As Integer, oCol As Integer, c As Integer, St As String, i As Integer
   Dim Gr As Control
   
   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   oRow = Gr.Row
   oCol = Gr.Col

   CurrFillStyle = Gr.FillStyle
   
   Gr.FillStyle = flexFillRepeat
   
   Gr.Row = Row
   Gr.RowSel = Row
   
   If EndCol < 0 Then
      EndCol = Gr.Cols - 1
   End If
   
   Gr.Col = StartCol

   Gr.ColSel = EndCol
   StyleType = UCase(StyleType)
   
   Do
      i = InStr(StyleType, "+")
      If i > 0 Then
         St = Left(StyleType, i - 1)
         StyleType = Mid(StyleType, i + 1)
      Else
         St = StyleType
      End If
      
      Select Case St
         Case "B": ' BOLD
            Gr.CellFontBold = True
         
         Case "I": ' ITALIC
            Gr.CellFontItalic = True
            
         Case "U": ' Underline
            Gr.CellFontUnderline = True
            
         Case "FC":   ' FORECOLOR
            Gr.CellForeColor = Style
            
         Case "BC":   ' BACKCOLOR
            Gr.CellBackColor = Style   'vb3DLight
            
         Case "T":    ' TEXT
            Gr.CellTextStyle = flexTextRaisedLight
            
         Case "ALIGN", "AL":   ' ALIGNMENT
            Gr.CellAlignment = Style
            
      End Select
   Loop While i > 0
   
   Gr.FillStyle = CurrFillStyle
   
   Gr.Row = oRow
   Gr.RowSel = oRow
   
   Gr.Col = oCol
   
   Gr.ColSel = oCol

End Sub

' Expande las columnas que tienen un blanco al final del título
' segun el ancho de la grillas
Public Sub FGrExpand(FGr As Control)
   Dim c As Integer, SWid As Integer, FWid As Integer
   Dim Fact As Single
   
   SWid = 0
   FWid = 0
   
   For c = 0 To FGr.Cols - 1
      If FGr.ColWidth(c) > 0 Then
         If Right(FGr.TextMatrix(0, c), 1) = " " Then
            FWid = FWid + FGr.ColWidth(c) + Screen.TwipsPerPixelX
         Else
            SWid = SWid + FGr.ColWidth(c) + Screen.TwipsPerPixelX
         End If
      End If
   Next c
   
   'If SWid + FWid + W.yScroll < FGr.Width Then
   If FWid > 0 And (FGr.Width - W.yScroll - SWid) > 0 Then
      Fact = (FGr.Width - W.yScroll - SWid) / FWid

      For c = 0 To FGr.Cols - 1
         If FGr.ColWidth(c) > 0 And Right(FGr.TextMatrix(0, c), 1) = " " Then
            FGr.ColWidth(c) = FGr.ColWidth(c) * Fact
         End If
      Next c

   End If

End Sub

Public Sub FGrFontBold(Grid As Control, ByVal Row As Integer, ByVal Col As Integer, ByVal Value As Boolean)
   Dim c As Integer, Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   If Row < 0 Then
      Row = Gr.rows - 1
   End If

   If Col >= 0 Then
      Gr.Row = Row
      Gr.Col = Col

      Gr.CellFontBold = Value
   Else
      For c = 0 To Gr.Cols - 1
         Gr.Row = Row
         Gr.Col = c
         Gr.CellFontBold = Value
      Next c
   End If

End Sub

Public Sub FGrBackColor(FGr As Control, ByVal Row As Integer, ByVal Col As Integer, ByVal Color As Long)
   Dim c As Integer, Gr As Control
   
   If TypeName(FGr) = "MSFlexGrid" Then
      Set Gr = FGr
   Else
      Set Gr = FGr.FlxGrid
   End If
   
   If Color = 0 Then
      Color = Gr.BackColorFixed
   End If

   If Row < 0 Then
      Row = Gr.rows - 1
   End If

   If Col >= 0 Then
      Gr.Row = Row
      Gr.Col = Col

      Gr.CellBackColor = Color
   Else
      For c = 0 To Gr.Cols - 1
         Gr.Row = Row
         Gr.Col = c
         Gr.CellBackColor = Color
      Next c
   End If

End Sub

Public Sub FGrForeColor(Grid As Control, ByVal Row As Integer, ByVal Col As Integer, ByVal Color As Long)
   Dim c As Integer, Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   If Col >= 0 Then
      Gr.Row = Row
      Gr.Col = Col
      Gr.CellForeColor = Color
   Else
      For c = 0 To Gr.Cols - 1
         Gr.Row = Row
         Gr.Col = c
         Gr.CellForeColor = Color
      Next c
   End If

End Sub

Public Sub FGrToolTip(Frm As Form, FGr As Control)
   Dim Tx As String
   Dim Col As Integer, Row As Integer
   
   On Error Resume Next
   If TypeName(FGr) = "MSFlexGrid" Then
      Col = FGr.MouseCol
      Row = FGr.MouseRow
   Else
      Col = FGr.FlxGrid.MouseCol
      Row = FGr.FlxGrid.MouseRow
   End If

   If Col < FGr.FixedCols Or Row < FGr.FixedRows Then
      FGr.ToolTipText = ""
      Exit Sub
   End If
   
   Tx = Trim(FGr.TextMatrix(Row, Col))

   If Tx <> "" And Frm.TextWidth(Tx) * 1.05 > FGr.ColWidth(Col) Then
      FGr.ToolTipText = Tx
   Else
      FGr.ToolTipText = ""
   End If
      
End Sub

Public Function FGrMouseCol(FGr As Control) As Integer

   On Error Resume Next
   If TypeName(FGr) = "MSFlexGrid" Then
      FGrMouseCol = FGr.MouseCol
   Else
      FGrMouseCol = FGr.FlxGrid.MouseCol
   End If

End Function
Public Function FGrMouseRow(FGr As Control) As Integer

   On Error Resume Next
   If TypeName(FGr) = "MSFlexGrid" Then
      FGrMouseRow = FGr.MouseRow
   Else
      FGrMouseRow = FGr.FlxGrid.MouseRow
   End If

End Function
Public Sub FGrToolTipText(FGr As Control, ByVal Text As String)

   On Error Resume Next
   If TypeName(FGr) = "MSFlexGrid" Then
      FGr.ToolTipText = Text
   Else
      FGr.FlxGrid.ToolTipText = Text
   End If

End Sub



Public Sub FGrEdPos(FGr As MSFlexGrid, Tx As TextBox)

   GrRow = FGr.Row
   GrCol = FGr.Col

   Tx.Left = FGr.Left + FGr.ColPos(GrCol) + 4 * Screen.TwipsPerPixelX
   Tx.Top = FGr.Top + FGr.RowPos(GrRow) + 4 * Screen.TwipsPerPixelY

   Tx.Width = FGr.CellWidth - Screen.TwipsPerPixelX
   Tx.Height = FGr.CellHeight - Screen.TwipsPerPixelY
   
   Tx.visible = True
   Tx.SetFocus

End Sub
Public Sub FGrClearPicture(Grid As Control, ByVal Row As Integer, ByVal Col As Integer)
   Dim oRow As Integer, oCol As Integer, Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   oRow = Gr.Row
   oCol = Gr.Col

   Gr.Row = Row
   Gr.Col = Col

   Set Gr.CellPicture = LoadPicture()
  
   Gr.Row = oRow
   Gr.Col = oCol

End Sub
Public Sub FGrSetPicture(Grid As Control, ByVal Row As Integer, ByVal Col As Integer, Pict As PictureBox, Optional bkcolor As Long = 0, Optional ByVal Align As AlignmentSettings = flexAlignCenterCenter)

   If Pict Is Nothing Then
      Call FGrClearPicture(Grid, Row, Col)
   Else
      Call FGrSetPicture1(Grid, Row, Col, Pict.Picture, bkcolor, Align)
   End If
   
End Sub
Public Sub FGrSetPicture1(Grid As Control, ByVal Row As Integer, ByVal Col As Integer, Pict As Picture, Optional bkcolor As Long = 0, Optional ByVal Align As AlignmentSettings = flexAlignCenterCenter)
   Dim oRow As Integer, oCol As Integer, Gr As Control

   If Pict Is Nothing Then
      Call FGrClearPicture(Grid, Row, Col)
      Exit Sub
   End If

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   oRow = Gr.Row
   oCol = Gr.Col

   Gr.Row = Row
   Gr.Col = Col

   Set Gr.CellPicture = Pict
   Gr.CellPictureAlignment = Align
   If bkcolor <> 0 Then
      Gr.CellBackColor = bkcolor
   End If
  
   Gr.Row = oRow
   Gr.Col = oCol

End Sub
Public Sub FGrSetImage(Grid As Control, ByVal Row As Integer, ByVal Col As Integer, Pict As Image, Optional bkcolor As Long = 0, Optional ByVal Align As AlignmentSettings = flexAlignCenterCenter)
   Dim oRow As Integer, oCol As Integer, Gr As Control

   If TypeName(Grid) = "MSFlexGrid" Then
      Set Gr = Grid
   Else
      Set Gr = Grid.FlxGrid
   End If

   oRow = Gr.Row
   oCol = Gr.Col

   Gr.Row = Row
   Gr.Col = Col

   Set Gr.CellPicture = Pict
   Gr.CellPictureAlignment = Align
   If bkcolor <> 0 Then
      Gr.CellBackColor = bkcolor
   End If
  
   Gr.Row = oRow
   Gr.Col = oCol

End Sub
' Configura la grilla de totales a partir de la grilla base
Public Sub FGrTotales(FGr As Control, FGrTot As Control, Optional bTopTot As Boolean = 1)
   Dim i As Integer

   FGrTot.Cols = FGr.Cols
   FGrTot.FixedCols = FGr.FixedCols
   FGrTot.Left = FGr.Left
   
   If FGr.ScrollBars = flexScrollBarVertical Or FGr.ScrollBars = flexScrollBarBoth Then
      FGrTot.Width = FGr.Width - W.xScroll
   Else
      FGrTot.Width = FGr.Width
   End If
   
   If bTopTot Then
      FGrTot.Top = FGr.Top + FGr.Height + 30
   End If
   
   For i = 0 To FGr.Cols - 1
      FGrTot.ColWidth(i) = FGr.ColWidth(i)

      If FGr.ColAlignment(i) = flexAlignCenterCenter Then
         FGrTot.ColAlignment(i) = flexAlignCenterTop
      ElseIf FGr.ColAlignment(i) = flexAlignRightCenter Then
         FGrTot.ColAlignment(i) = flexAlignRightTop
      Else ' If FGr.ColAlignment(i) = flexAlignLeftCenter Then
         FGrTot.ColAlignment(i) = flexAlignLeftTop
      End If
      
   Next i

   FGrTot.RowHeight(0) = 270
   FGrTot.Height = FGrTot.RowHeight(0) + 10
   
   If TypeName(FGrTot) = "MSFlexGrid" Then
      FGrTot.BackColorBkg = vbWindowBackground
   Else
      FGrTot.FlxGrid.BackColorBkg = vbWindowBackground
   End If
   

End Sub

Public Sub FGrSetRowVisible(FGr As Control, ByVal Row As Long)
   Dim hRow As Integer, nRows As Integer

   On Error Resume Next
   
   If TypeName(FGr) = "MSFlexGrid" Then
      Set FGr = FGr
   Else
      Set FGr = FGr.FlxGrid
   End If
   
   If Row < FGr.TopRow Then
      FGr.TopRow = Row
   Else

      hRow = FGr.RowHeight(Row)
      If hRow Then
         nRows = Int(FGr.Height / (hRow + Screen.TwipsPerPixelY))
      Else
         nRows = 0
      End If
      
      If Row >= FGr.TopRow + nRows - FGr.FixedRows Then
         FGr.TopRow = Row
      End If

   End If

End Sub

Public Function FGrBeforeDate(FGr As Control, ByVal Row As Integer, ByVal ColhFecha As Integer, ByVal DefDat As Long) As String
   Dim Dt As Long

   Dt = Val(FGr.TextMatrix(Row, ColhFecha))

   If Dt <= 0 Then
      FGrBeforeDate = Format(DefDat, EDATEFMT)
   Else
      FGrBeforeDate = Format(Dt, EDATEFMT)
   End If

End Function

Public Function FGrAcceptDate(FGr As Control, ByVal Row As Integer, ByVal ColhFecha As Integer, ByVal Value As String) As String
   Dim Dt As Long

   If Trim(Value) = "" Then
      FGr.TextMatrix(Row, ColhFecha) = ""
      Value = ""
   Else
      Dt = GetDate(Value)
      If Dt <= 0 Then
         MsgBox1 "Fecha inválida", vbExclamation
         Value = ""
      Else
         FGr.TextMatrix(Row, ColhFecha) = Dt
         Value = Format(Dt, DATEFMT)
      End If
   End If

   FGrAcceptDate = Value

End Function

Public Function FGrChkMaxSize(FGr As Control, Optional ByVal bMsg As Boolean = 1) As Boolean

   If FGr.rows * FGr.Cols > 300000 Then  ' 350000
      FGrChkMaxSize = True
      
      If bMsg Then
         MsgBox1 "Esta consulta supera la cantidad máxima de celdas soportada por este objeto, segun las columnas que tiene, soporta " & Round(300000 / FGr.Cols) - 2 & " Filas" & vbCrLf & "Si es posible aplique un filtro y consulte nuevamente.", vbExclamation
      End If
   Else
      FGrChkMaxSize = False
   End If

End Function

Public Sub FGrMergeCols(FGr As Control, ByVal Row As Integer, ByVal ColIni As Integer, ByVal ColFin As Integer, ByVal Text As String)
   Dim c As Integer, Gr As Control ' MSFlexGrid

   If TypeName(FGr) = "MSFlexGrid" Then
      Set Gr = FGr
   Else
      Set Gr = FGr.FlxGrid
   End If

   Gr.Row = Row
   Gr.Col = ColIni
   Gr.MergeCells = flexMergeRestrictRows
   
   For c = ColIni To ColFin
   
      Gr.TextMatrix(Row, c) = Text
      Gr.MergeCol(c) = True
   Next c
   
   Gr.MergeRow(Row) = True
'   Gr.MergeCells = flexMergeNever

End Sub

Public Function FGrColsWid(FGr As Control) As Long
   Dim i As Integer, W As Long, Gr As Control

   If TypeName(FGr) = "MSFlexGrid" Then
      Set Gr = FGr
   Else
      Set Gr = FGr.FlxGrid
   End If

   W = 0
   For i = 0 To Gr.Cols - 1
      W = W + Gr.ColWidth(i)
   Next i
   
   FGrColsWid = W

End Function

Public Function FGrWidth(FGr As Control) As Long
   Dim Gr As Control, c As Integer, wID As Long
   
   If TypeName(FGr) = "MSFlexGrid" Then
      Set Gr = FGr
   Else
      Set Gr = FGr.FlxGrid
   End If

   wID = 0
   For c = 0 To Gr.Cols - 1
      wID = wID + Gr.ColWidth(c)
      If Gr.ColWidth(c) > 0 Then
         wID = wID + Screen.TwipsPerPixelX
      End If
   Next c

   wID = wID + Screen.TwipsPerPixelX

   FGrWidth = Int((wID + W.yScroll) / 60 + 0.9) * 60

End Function
' Para revisar lo básico en Grid_BeforeEdit
Public Function FgrChkBefore(FGr As Control, ByVal ColOblig As Integer, ByVal Row As Integer, ByVal Col As Integer) As Boolean

   If Row < FGr.FixedRows Or Col < FGr.FixedCols Then
      Exit Function
   End If

   If FGr.TextMatrix(Row - 1, ColOblig) = "" Then
      Exit Function
   End If

   If Col > FGr.FixedCols Then
      If FGr.TextMatrix(Row, Col - 1) = "" Then
         Exit Function
      End If
   End If

   FgrChkBefore = True

End Function
' 29 ago 2018: para mantener una cierta cantidad de lineas al final
Public Sub FGrEdRows(Gr As Control, ByVal Row As Integer, Optional ByVal NewRows As Integer = 5)
   Dim rd As Boolean

   If Row >= Gr.rows - 1 Then
      rd = Gr.Redraw
      Gr.Redraw = False
      Gr.rows = Gr.rows + NewRows
      Gr.Redraw = rd
   End If

End Sub
Public Sub FGrLeftCol(GrTot As Control, Gr As Control)
   Dim i As Integer
   For i = 0 To Gr.Cols - 1
      GrTot.ColWidth(i) = Gr.ColWidth(i)
   Next i
   
   GrTot.LeftCol = Gr.LeftCol

End Sub
' 25 oct 2019: elimina filas entre Grid.row y Grid.RowSel
Public Function FGrDelMultiRows(Grid As Control, ByVal ColControl As Integer, Optional ByVal idCol As Integer = -1, Optional ByVal StCol As Integer = -1) As Integer
   Dim n As Integer, i As Integer, Ini As Integer, Fin As Integer

   Ini = Min(Grid.Row, Grid.RowSel)
   Fin = Max(Grid.Row, Grid.RowSel)
   
   If idCol = -1 Then
      idCol = Grid.Cols - 2
   End If
   
   If StCol = -1 Then
      StCol = Grid.Cols - 1
   End If
   
   Grid.Redraw = False
   
   For i = Ini To Fin
      
      If Grid.TextMatrix(i, ColControl) <> "" Then
         n = n + 1
         Call FGrModRow(Grid, i, FGR_D, idCol, StCol)
         
         Grid.rows = Grid.rows + 1
      End If
   Next i
   Grid.Redraw = True

   FGrDelMultiRows = n
   
End Function

Public Function FGrSelRows(Grid As Control)
   FGrSelRows = Max(Grid.Row, Grid.RowSel) - Min(Grid.Row, Grid.RowSel) + 1
End Function

Public Function FgrFixedText(Grid As Control, ByVal Row As Integer) As String
   Dim c As Integer, Buf As String
   
   For c = 0 To Grid.FixedCols
      If Grid.ColWidth(c) > 0 Then
         Buf = Buf & " " & Grid.TextMatrix(Row, c)
      End If
   Next c
   
   FgrFixedText = Mid(Buf, 2)
   
End Function
#If DATACON > 0 Then
'LLena una grilla como una fillcombo, sólo con dos datos
Public Function FillGrLista(Qry As String, Gr As Control, ByVal LASTCOL As Integer, wCol() As Byte)
   Dim Rs As Recordset
   Dim i As Integer, c As Integer
   
   Set Rs = OpenRs(DbMain, Qry)
   
   i = Gr.FixedRows
   Gr.Redraw = False
   Gr.rows = i
   
   For c = 0 To LASTCOL
'      wCol(c) = Rs(c).Size
      wCol(c) = FldSize(Rs(c))
   Next c
   
   Do While Rs.EOF = False
      Gr.rows = i + 1
      
      For c = 0 To LASTCOL
         If IsNull(Rs(c)) = False Then
            Gr.TextMatrix(i, c) = vFld(Rs(c))
         End If
      Next c
      
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   Call FGrVRows(Gr)
   
   Gr.Redraw = True
  
End Function
#End If '  DATACON

Public Function FGrInfo(Gr As Control) As String
   Dim Txt As String, i As Integer
   
   For i = 0 To Gr.FixedRows - 1
      Txt = Txt & " " & Gr.TextMatrix(i, Gr.Col)
   Next i

   FGrInfo = Trim(Txt) & ": " & Gr.Text

End Function

'2861570
Public Function FGr2Clip_membr(Grid As Control, Optional ByVal Title As String = "", Optional ByVal bIncludeCero As Boolean = False, Optional ByVal bNoQuotes As Boolean) As Long
   Dim Clip As String
   Dim Membrete As String
   
   Clip = FGr2String(Grid, Title, bIncludeCero, , False, bNoQuotes)
    If MsgBox1("¿Desea agregar datos básicos de la empresa (rut, nombre, dirección giro, rep. Legal)?.", vbInformation + vbYesNo) = vbYes Then
     Membrete = "Razón Social " & vbTab & gEmpresa.RazonSocial & vbCrLf
      Membrete = Membrete & " Rut " & vbTab & gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut) & vbCrLf
      Membrete = Membrete & " Dirección " & vbTab & gEmpresa.Direccion & ", " & IIf(gEmpresa.Ciudad <> "", FCase(gEmpresa.Ciudad), FCase(gEmpresa.Comuna)) & vbCrLf
      Membrete = Membrete & " Giro " & vbTab & gEmpresa.Giro & vbCrLf
      Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal1 & vbCrLf
      If gEmpresa.RutRepLegal1 <> "" Then
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & "-" & DV_Rut(gEmpresa.RutRepLegal1) & vbCrLf
      Else
      Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal1 & vbCrLf
      End If
      
      If gEmpresa.RepConjunta Then
        Membrete = Membrete & " Rep. Legal " & vbTab & gEmpresa.RepLegal2 & vbCrLf
        If gEmpresa.RutRepLegal2 <> "" Then
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & "-" & DV_Rut(gEmpresa.RutRepLegal2) & vbCrLf & vbCrLf
        Else
        Membrete = Membrete & " Rut Rep. Legal " & vbTab & gEmpresa.RutRepLegal2 & vbCrLf & vbCrLf
        End If
      End If
      
      Clip = Membrete & Clip
      End If
      
      FGr2Clip_membr = SetClipText(Clip)

End Function
'2861570
