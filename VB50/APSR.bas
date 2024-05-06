Attribute VB_Name = "APSR"
Option Explicit
Global Const GE_NADA = 0
Global Const GE_NUEVO = 1
Global Const GE_MODIFICA = 2
Global Const GE_ELIMINA = 3
Global Const GE_EXISTE = 4

'Color
Public Const COLOR_BUTONFACE = &H8000000F

Public gYearNow As Long
Public gMonthNow As Integer
Public gDayNow As Integer
Public Const EDIAMESFMT = "dd\/mm"

'Nombre PC en el Ini
Public gNamePcIni As String

'EQUIVALENCIAS
Public Const TW_ML = 56.7
Public Const TW_CM = 567

Type ConfigPrt_t
   FontName           As String
   FontBold           As Boolean
   FontSize           As Single
   AlignTop           As Single
   MargenLeft         As Single
End Type

Type ComDlg_t
   Filter      As String
   DialogTitle As String
   FileTitle   As String
   InitDir     As String
   Filename    As String
   Flags       As Long
End Type
'Public Sub FillComboDate(Cmb As Control, Db As Database, Qry As String, idSel As Long)
'   Dim Rs As Recordset
'
'   Set Rs = OpenRs(Db, Qry)
'   If Rs Is Nothing Then
'      Exit Sub
'   End If
'
'   Do Until Rs.EOF
'      Cmb.AddItem FmtDate(vFld(Rs(0)))
'      Cmb.ItemData(Cmb.NewIndex) = Val(vFld(Rs(1)))
'
'      If idSel >= 0 And idSel = vFld(Rs(1)) Then
'         Cmb.ListIndex = 0
'      End If
'
'      Rs.MoveNext
'   Loop
'
'   Call CloseRs(Rs)
'
'   If idSel = -1 And Cmb.ListCount > 0 Then
'      Cmb.ListIndex = 0
'   End If
'
'End Sub

'Busca en un path el nombre de un archivo
Public Function FindFile(PathFile As String) As String
   Dim Buf As String
   Dim i As Integer, j As Integer, IniName As Integer
   
   i = Len(PathFile)
   
   For j = i - 1 To 1 Step -1
      Buf = Mid(PathFile, j, 1)
      
      If Buf = "\" Then
         IniName = j + 1
         j = 1
         
      End If
      
   Next j
      
   FindFile = Mid(PathFile, IniName)
   
End Function

Private Sub SizeImage(picBox As PictureBox, sizePic As Picture, sizeWidth As Single, sizeHeight As Single)
  Screen.MousePointer = vbHourglass

On Error GoTo ErrorSize
  picBox.Picture = LoadPicture("")
  picBox.Width = sizeWidth
  picBox.Height = sizeHeight
  picBox.AutoRedraw = True
  picBox.PaintPicture sizePic, 0, 0, sizeWidth, sizeHeight
  picBox.Picture = picBox.image
  picBox.AutoRedraw = False
  GoTo EndSize
  
ErrorSize:
 ' lblStatus = ""
  
EndSize:
  Screen.MousePointer = vbDefault
  
End Sub
Public Function RutGotFocus(Rut As String) As String
   Dim Dig As String
   
        Debug.Print "*** Usar RUT_GotFocus ***"

   If Trim(Rut) <> "" Then
      Dig = Right(Rut, 2)
      Rut = vFmtRut(Rut)
      Rut = Rut & Dig
      
   End If
   RutGotFocus = Rut
   
End Function
Public Function RutLostFocus(Rut As Control) As String

        Debug.Print "*** Usar RUT_LostFocus ***"

   If Trim(Rut) <> "" Then
      If Not MsgValidRut(Rut) Then
         RutLostFocus = Rut
         Rut.SetFocus
         Exit Function
      
      End If
      RutLostFocus = FmtRut(vFmtRut(Rut))
      
  End If
  
End Function
#If DATACON = DAO_CON Then
Public Sub FillCbDate(Cmb As Control, Db As Database, Qry As String, idSel As Long)
#Else
Public Sub FillCbDate(Cmb As Control, Db As Connection, Qry As String, idSel As Long)
#End If
   Dim Rs As Recordset

   Set Rs = OpenRs(Db, Qry)
   If Rs Is Nothing Then
      Exit Sub
   End If
   
   Do Until Rs.EOF
      Cmb.AddItem FmtDate(vFld(Rs(0)))
      Cmb.ItemData(Cmb.NewIndex) = Val(vFld(Rs(1)))
      
      If idSel >= 0 And idSel = vFld(Rs(1)) Then
         Cmb.ListIndex = 0
      End If

      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   If idSel = -1 And Cmb.ListCount > 0 Then
      Cmb.ListIndex = 0
   End If

End Sub
Public Function SelText(Cb As Control, ByVal str As String) As Integer
   SelText = CbSelText(Cb, str)
End Function

Public Function FillCbAno(cbAno As ComboBox, Optional Sel As Integer = -1, Optional ByVal AnosAntes As Integer = 5)
   Dim Desde As Integer, Hasta As Integer, Ano As Integer
   
   Hasta = Year(Now) + 1
   Desde = Hasta - AnosAntes
   
   If Sel > Hasta Then
      Hasta = Sel
   End If
   
   For Ano = Hasta To Desde Step -1
   
      Call AddItem(cbAno, Ano, Ano, Ano = Sel)
               
   Next Ano
   
End Function
''Llena combo años , con el año actual y pasado
'Public Sub FillAno(Ano As ComboBox)
'
'   Ano.AddItem Year(Now) - 1
'   Ano.AddItem Year(Now)
'   Ano.AddItem Year(Now) + 1
'   Ano.ListIndex = 0
'
'End Sub
Sub ApsrInit()
   gYearNow = Year(Int(Now))
   gMonthNow = Month(Int(Now))
   gDayNow = Day(Int(Now))
   
End Sub
Public Function FindItemcb(Cb As ComboBox, Texto As String) As Boolean
   Dim i As Integer
   
   FindItemcb = False
   For i = 0 To Cb.ListCount - 1
      If Cb.List(i) = Texto Then
         Cb.ListIndex = i
         i = Cb.ListCount - 1
         FindItemcb = True
         
      End If
      
   Next i

End Function
Public Function SelTxt(Cb As Control, ByVal str As String) As Integer
   Dim i As Integer

   For i = 0 To Cb.ListCount - 1
      If Cb.List(i) = Trim(str) Then
         SelTxt = i
         Exit Function
         
      End If
      
   Next i

   SelTxt = -1
   
End Function
'Llena una combo o list con la lista de impresoras que estan definidas
Public Function FillCbPrinter(Cb As Control)
    Dim NamePrinter As Printer
    
    For Each NamePrinter In Printers
      Cb.AddItem NamePrinter.DeviceName
    Next
        
    If Cb.ListIndex < 0 And Cb.ListCount > 0 Then
      Cb.ListIndex = 0
      
   End If
End Function
'Setea una impresora definida por defecto
Public Function SetPrinterDefault(sDeviceName As String) As Boolean
    Dim NamePrinter As Printer
    Dim bool As Boolean
   
    bool = False
    For Each NamePrinter In Printers
        If NamePrinter.DeviceName = sDeviceName Then
            Set Printer = NamePrinter
            bool = True
            Exit For
        End If
    Next
    SetPrinterDefault = bool
    
End Function

Public Function FillTotcbAno(cbAno As ComboBox, TotAnt As Integer)
   Dim Ano As Long
   Dim i As Integer
   
   Ano = Year(Now) - TotAnt
   
   For i = 1 To TotAnt + 4
      cbAno.AddItem Ano
      Ano = Ano + 1
      
   Next i
   
End Function
Public Function KillFile(NameFile As String) As Boolean
   On Error Resume Next
   
   KillFile = False
   Kill (NameFile)
   If Err = 70 Or (Dir(NameFile) <> "") Then
      Exit Function 'No se pudo elliminar
   End If
   KillFile = True
   
End Function
'Ejemplo Filter:"Todas (*.jpg,*.gif,*.dib,*.bmp)|*.jpg;*.gif;*.dib;*.bmp|Imágenes JPG(*.jpg)|*.jpg|Imágenes GIF(*.gif)|*.gif|Imágenes DIB (*.dib)|*.dib|Imágenes BMP (*.bmp)|*.bmp"
'FileTitle =Nombre Archivo
Public Function FileNameComDlg(Cm_ComDlg As CommonDialog, Comdlg As ComDlg_t) ' Filter As String, DialogTitle As String, Optional FileTitle As String = "") As String
   FileNameComDlg = ""
   
   Cm_ComDlg.CancelError = True
   Cm_ComDlg.Filename = Comdlg.Filename
   Cm_ComDlg.InitDir = Comdlg.InitDir
   Cm_ComDlg.Filter = Comdlg.Filter
   Cm_ComDlg.DialogTitle = Comdlg.DialogTitle
   
   If Comdlg.Flags = 0 Then
      Cm_ComDlg.Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
   Else
      Cm_ComDlg.Flags = Comdlg.Flags
   End If
   
   On Error Resume Next
   Cm_ComDlg.Action = 1
   
   If Err = cdlCancel Then
      Exit Function
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & Cm_ComDlg.Filename, vbExclamation
      Exit Function
   End If

   If Cm_ComDlg.Filename = "" Then
      Exit Function
   End If
   Err.Clear
   
   Comdlg.FileTitle = Cm_ComDlg.FileTitle
   FileNameComDlg = Cm_ComDlg.Filename
      
   Err.Clear
   
End Function
Public Function CheckConfigPrt() As Boolean
   Dim DeviceName As String
   
   CheckConfigPrt = False
   
   On Error Resume Next
   
   DeviceName = Printer.DeviceName
   If Err = 484 Or Err = 482 Then
      MsgBox1 "Error, no existe ninguna impresora configurada en su equipo", vbExclamation
      Exit Function
      
   ElseIf Err Then
      MsgBox1 Err
      Exit Function
      
   ElseIf DeviceName = "" Then
      MsgBox1 "Error, no existe ninguna impresora configurada en su equipo", vbExclamation
      Exit Function
      
   End If
   
   CheckConfigPrt = True
   
End Function
'Agregué 27/10/2005 el Byval porque no aparecia un texto, porque lo predia sin éste
Public Sub AjusteTexto(ByVal AuxTexto As String, TLeft As Integer, RightX As Long)
   Dim i As Integer, j As Integer
   Dim k As Integer, b As Integer
   
   Printer.Print Tab(TLeft);
   AuxTexto = AuxTexto & " "
   Do While AuxTexto <> ""
      
      If Printer.TextWidth(AuxTexto) <= RightX Then
         Printer.Print Tab(TLeft); AuxTexto
         j = Len(AuxTexto) + 1
         
      Else
      
         k = 0
         b = 0
         For i = Len(AuxTexto) To 1 Step -1
            If Printer.TextWidth(Left(AuxTexto, i)) <= RightX Then
            
               k = InStr(Left(AuxTexto, i), vbLf)
               If k Then
                  b = 1
                  If Mid(AuxTexto, k - 1, 1) = vbCr Then
                     b = b + 1
                  End If
   
                  j = k - b
               Else
                  j = i
                  Do While Mid(AuxTexto, j, 1) <> " "
                     j = j - 1
                  Loop
               
               End If
               
               Printer.Print Tab(TLeft); Left(AuxTexto, j)
               Exit For
            End If
            
         Next i
         
      End If
      
      If k Then
         j = j + b
         
         Printer.Print Tab(TLeft);
      End If
     
      AuxTexto = Trim(Mid(AuxTexto, j + 1))

   Loop
   
End Sub
Public Sub NamePcInscrip()
   Dim Rc As Integer
   
   If Trim(gNamePcIni) <> "" And gAppCode.Demo Then
      If gNamePcIni <> w.PcName Then
         'Estuvo inscrito y cambio Nombre de PC
         MsgBox1 "¡ATENCION! Usted ha cambiado el nombre de PC por " & w.PcName & "." & vbNewLine & gAppCode.Title & " esta inscrito con el nombre de PC " & gNamePcIni & "." & vbNewLine & "Para que el sistema no quede en modo DEMO, vuelva a poner el nombre de PC con el que inscribio " & gAppCode.Title & ".", vbExclamation
      End If
      
   End If
   
End Sub


