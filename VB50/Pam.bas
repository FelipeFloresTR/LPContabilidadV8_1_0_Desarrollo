Attribute VB_Name = "PAM"
'
' Funciones Generales
'
Option Explicit

' Simbolos de datos
' %  Entero
' &  Long
' !  Single
' #  Double
' $  String


Public gFrmMain As Form
Public gPrinter As Printer
#If K_SelPrinter Then
Public gPrtDlg  As CommonDialog
#End If

Public Const PI As Double = 3.14159265359
Public Const Euler As Double = 2.71828182845905 ' e

' Formatos: permite hasta 4 secciones
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications
' One section only: The format expression applies to all values.
' Two sections: The first section applies to positive values and zeros, the second to negative values.
' Three sections: The first section applies to positive values, the second to negative values, and the third to zeros.
' Four sections: The first section applies to positive values, the second to negative values, the third to zeros, and the fourth to Null values.

Public Const NUMFMT As String = "#,##0"
Public Const NUMFMTN As String = "#,##0;(#,##0)"
Public Const DBLFMT1 As String = "#,##0.0"
Public Const DBLFMT2 As String = "#,##0.00"
Public Const DBLFMT3 As String = "#,##0.000"
Public Const DBLFMT4 As String = "#,##0.0000"
Public Const DBLFMT5 As String = "#,##0.00000"
Public Const DBLFMT6 As String = "#,##0.000000"
Public Const DBLFMT  As String = DBLFMT1    ' por compatibilidad con código antiguo

Public Const BL0_DBLFMT1 As String = "#,###.0;#,###.0;#,###.#"
Public Const BL0_DBLFMT2 As String = "#,###.00;#,###.00;#,###.##"

Public Const DBLFMT1DO As String = "#,##0.#"    ' 1 decimal, si es que tiene
Public Const DBLFMT2DO As String = "#,##0.##"   ' 2 decimales, si es que tiene

Private lSigDec As String * 1  ' signo decimal: , o . para KeyDec

Public Const DATEFMT2 As String = "dd mmm yy"
Public Const EDATEFMT2 As String = "dd\/mm\/yy"
Public Const DATEFMT3 As String = "yyyy-mm-dd"  ' Para copiar a excel

Public Const MONTHFMT As String = "mmm yyyy"
Public Const EMONTHFMT As String = "mm\/yyyy"    ' Para SelMonth en calendar

Public Const KEY_MENU95 = 93 ' Tecla de Menú - Windows

' Para funciones que no retornan Errores
Public gFwErr As Long
Public gFwError As String

Public gMsgBoxLog As Boolean  ' para que Msgbox1 grabe en el Log y que no muestre mensaje

' Colores del sistema
' vbWindowBackground
Public Const VBCOLOR_BUTTONFACE = &H8000000F
Public Const VBCOLOR_WINDOWTEXT = &H80000008
Public Const VBCOLOR_HIGHLIGHT = &H8000000D
Public Const VBCOLOR_HIGHLIGHTTEXT = &H8000000E

Public Const VBCOLOR_LIGHTYELLOW = &HC0FFFF        ' suave
Public Const VBCOLOR_LIGHTYELLOW2 = &H80FFFF

Public Const VBCOLOR_LIGHTGREEN = &HC0FFC0
Public Const VBCOLOR_LIGHTRED = &HE6E8FD
Public Const VBCOLOR_LIGHTRED2 = &HC0C0FF

Public Const GRAY = &HC0C0C0
Public Const DK_GREEN = &HC000&

Public Const BKCOLOR_MOD = VBCOLOR_LIGHTYELLOW    ' para marcar que algo cambió

Public gBkColOblig  As Long   ' Para campos obligatorios

Public Const FRMALIGN = vbCenter
Public Const SRCCOPY = &HCC0020

Public Const EM_SETREADONLY = &H41F

Public Const TWIPS_INCH = 1440   ' Inch = Twips / TWIPS_INCH
Public Const TWIPS_CM = 567      ' cm = Twips / TWIPS_CM

' Si Año < 90 se asume 2000 + Año
Public Const AÑO_2DIG = 90

' Tipo de base utilizada, para algunas funciones
Public Const PDB_ACCESS = 0
Public Const PDB_ODBC = 1
Public gPamDbTipo As Byte

Public gUpCase    As Integer
Public gAppPath   As String   ' igual que App.path pero nunca termina en \
'Public AppPath    As String   ' igual que App.path pero nunca termina en \
Public gTmpDir    As String
'Public gUserName  As String
'Public gPCName    As String

Public gValidRut  As Integer

Public gGrayText As Long  'Color Grayed Text

Public NL As String  ' NewLine          0x0A
Public CR As String  ' Carriare Return  0x0D
Public Tb As String  ' Tab
Public CRNL As String

Public gNullWords As String
Public gLoWords As String   ' Para la función FCase
Public gSkWords As String   ' Para la función FCase

Public gSiNo(2) As String
Public Const VAL_NO = 0
Public Const VAL_SI = 1
Public Const VAL_OTRO = 2

Type W_Info
   YCaption    As Integer  ' Alto del caption
   yMenu       As Integer  ' Alto del menu
   xFrame      As Integer  ' Ancho del borde de la ventana
   yFrame      As Integer  ' Alto del borde de la ventana
   xScroll     As Integer  ' Ancho de la barra de scroll vertical
   yScroll     As Integer  ' Alto de la barra de scroll horizontal
   dx          As Single
   dy          As Single
   
   WinDir      As String   ' Directorio de Windows
   WinDrv      As String   ' C: o D:
   
   PcName      As String
   UserName    As String
   Mac         As String
   
   TmpDir      As String
   DownDir     As String

   AppPath     As String
   Version     As String
   FVersion    As Long     ' Fecha de la versión
   FullVer     As String
   FStart      As Double   ' Fecha en que se inicio el programa
   
   NumDecSym   As String   ' Simbolo decimal en números
   NumSepSym   As String   ' Separador de miles en números
   CurDecSym   As String   ' Simbolo decimal en monedas
   CurSepSym   As String   ' Separador de miles en monedas
   ShortDate   As String
   DateSep     As String
   
   DefMailer   As String
      
   InDesign    As Boolean
End Type
Public W As W_Info

Public gNomMes(12) As String
Public gMonNam(12) As String
Public gDiaSem(7) As String
Public gDiaSem1(7) As String  ' Para Weekday con vbMonday, 1: lunes

Private lRndInit As Boolean    ' se inicio el Randomize en PamRandomize

Type FindApp_t
   hWnd As Long
   Caption As String
End Type
Private lFindApp As FindApp_t

Type UnixFile_t
   Fd    As Long
   Buf   As String
End Type

' Estructura general para retornar valores desde Forms
Type RC_t
   Rc       As Long
   Buf      As String
   Value    As Long
   DValue   As Double
   iValue   As Integer
   SValue   As String
End Type
Public gRc As RC_t

Type ExtInfo_t
   bReg        As Boolean  ' Se obtuvo del Registry

   Ext         As String
   Path        As String
   
   OpenCmd     As String
   PrintCmd    As String
   
   UseDde      As Boolean
   
   DdeOpen     As String
   DdeOpenApp  As String
   DdeOpenTopic As String
   
   DdePrint    As String
   DdePrtApp   As String
   DdePrtTopic As String

End Type

' Para la funcion FitPicture
Type InfoPict_t
   fname    As String
   Width    As Integer
   Height   As Integer
End Type


' Transferencia de caracteres
Public ChrMask(5, 1) As String



Public Function ActivateApp(AppName As String, Optional ByVal nCmdShow As Long = vbMaximized) As Boolean
   Dim hWord As Long
   
   hWord = FindApp(AppName)

   If hWord Then
      
      Call ShowWindow(hWord, nCmdShow)
      Call SetWindowPos(hWord, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
      Call ApiSetFocus(hWord)
      ActivateApp = True
   Else
      Beep
      ActivateApp = False
   End If

End Function

'
' Esta funcion reemplaza los ' por chr(255) o al revez, para
' los comandos SQL
'
' Bool:  True  => cambia ' por `
'        False => cambia ` por '
'
'
Function charSQL(ByVal Buf As String, ByVal bool As Integer) As String
   Dim i As Integer, j As Integer
   Dim ChFrom As String, ChTo As String

   If bool Then
      ChFrom = "'"
      ChTo = "`"
   Else
      ChFrom = "`"
      ChTo = "'"
   End If

   j = 1
   Do
      i = InStr(j, Buf, ChFrom)
      If i = 0 Then
         Exit Do
      End If

      j = j + i
      Mid(Buf, j - 1, 1) = ChTo

   Loop
      
   charSQL = Buf

End Function

'
' Para el WHERE en una consulta SQL con campos nulos
'
Function ChkIsNull(ByVal Value As String, ByVal isNum As Integer) As String

   If Value = "" Then
      ChkIsNull = "Is Null"
   ElseIf isNum Then
      ChkIsNull = "= " & Value
   Else
      ChkIsNull = "= '" & Value & "'"
   End If
        
End Function

'
' Para UPDATE en un query SQL
'
Function ChkNull(ByVal Value As String, ByVal isNum As Integer) As String
   If Value = "" Then
      ChkNull = "Null"
   ElseIf isNum Then
      ChkNull = Value
   Else
      ChkNull = "'" & Value & "'"
   End If

End Function
'
' Hola, 12, "pero"
Function CsvId(Buf As String) As String
   Dim i As Integer, j As Integer
    
   i = InStr(Buf, ",")
   
   If i Then
      CsvId = Trim(Left(Buf, i - 1))
   
      Buf = Trim(Mid(Buf, i + 1))
   
   Else
      CsvId = Trim(Buf)
      Buf = ""
   End If


End Function
'
' "Hola", 12, "pero"
Function CsvStr(Buf As String) As String
   Dim i As Integer, j As Integer
    
   i = InStr(Buf, """")
   j = InStr(i + 1, Buf, """")

   CsvStr = Trim(Mid(Buf, i + 1, j - i - 1))
   
   i = InStr(j + 1, Buf, ",")
   If i Then
      Buf = Trim(Mid(Buf, i + 1))
   Else
      Buf = ""
   End If

End Function
'
' Hola<tab>12<tab>pero
' Es igual a NextField
Function GetBufStr(Buf As String) As String
   Dim i As Integer, j As Integer
    
   i = InStr(Buf, vbTab)
   
   If i Then
      GetBufStr = Trim(Left(Buf, i - 1))
   
      Buf = Trim(Mid(Buf, i + 1))
   Else
      GetBufStr = Trim(Buf)
      Buf = ""
   End If

End Function

Function DateVal(ByVal Fecha As String) As Long
   If Len(Fecha) < 10 Then
      DateVal = 0
      Exit Function
   End If

   DateVal = DateSerial(Val(Mid(Fecha, 1, 4)), Val(Mid(Fecha, 6, 2)), Val(Mid(Fecha, 9, 2)))

End Function

Function DeSQL(ByVal Buf As String) As String
   Dim i As Integer

   ' Para compatibilidad hacia atrás
   Buf = ReplaceStr(Buf, Chr(253), "'")        ' reemplaza 253 por '
   Buf = ReplaceStr(Buf, Chr(254), Chr(34))   ' reemplaza 252 por "
   Buf = ReplaceStr(Buf, Chr(255), "|")        ' reemplaza 251 por |
   Buf = ReplaceStr(Buf, Chr(248), Chr(10))   ' reemplaza 250 por nl
   Buf = ReplaceStr(Buf, Chr(247), Chr(13))   ' reemplaza 249 por cr

   For i = 0 To 5
      Buf = ReplaceStr(Buf, ChrMask(i, 1), ChrMask(i, 0))
   Next i

   DeSQL = Trim(Buf)

End Function

'
'  Despliega un Menu Popup donde esta el cursor
'
Sub DispPopup(Frm As Form, Menu As Control, ByVal nPos As Integer)
   Dim HMenu0 As Long, hMenu1 As Long
   Dim hMenu As Long
   Dim Rc As Integer
   Dim Pt As POINTAPI_T
   Dim nBar As Integer

   If gFrmMain Is Nothing Then
      Debug.Print "*** Falta asignar el gFrmMain ***"
      Exit Sub
   End If

   HMenu0 = GetMenu(gFrmMain.hWnd)

   If Frm.MDIChild Then
      If Frm.WindowState = vbNormal Then
         nBar = 0
      Else
         nBar = 1
      End If
   Else
      nBar = 0
   End If
   
   Menu.visible = True
   Menu.Enabled = True

   hMenu1 = GetSubMenu(HMenu0, nBar)

   If nPos >= 0 Then

      hMenu = GetSubMenu(hMenu1, nPos) ' posición del submenu
      If hMenu = 0 Then
         MsgBeep vbExclamation ' El submenu no existe o no es el correcto
         Exit Sub
      End If

   Else
      hMenu = hMenu1

   End If

   Call GetCursorPos(Pt)

   ' Rc = TrackPopupMenu(HMenu, 2, Pt.X, Pt.Y, 0, Frm.hWnd, 0)
   Rc = TrackPopupMenu(hMenu, 2, Pt.x, Pt.Y, 0, gFrmMain.hWnd, 0)
   DoEvents
   
   'On Error Resume Next
   Menu.visible = False
   Menu.Enabled = False

End Sub

Public Sub EnableFrame(Frm As Form, fr As Frame, ByVal bEnable As Boolean)
   Dim i As Integer, CmdEsp As String, bSkip As Boolean
   Dim TName As String, CName As String
   Dim Ctrl As Control, p As Control
   Static MsgNotSup As String

   ' Estos botones no los modifica
   CmdEsp = ",bt_cancel,bt_cancelar,bt_close,bt_cerrar,bt_print,bt_imprimir,bt_preview,bt_copyexcel,"

   On Error Resume Next

   For i = 0 To Frm.Controls.Count - 1

      Set Ctrl = Frm.Controls(i)
      Set p = Ctrl
      
      Debug.Print p.Name
'      If P.Name = "Cb_Moneda" Then
'         Beep
'      End If
      
      bSkip = False
      Do
         Debug.Print p.Container.Name
         If p.Container.hWnd = Frm.hWnd Then
            bSkip = True
            Exit Do
         End If
      

         If p.Container.hWnd = fr.hWnd Then
            Exit Do
         End If
      
'         If P.Parent.hWnd = Frm.hWnd Then
'            bSkip = True
'            Exit Do
'         End If
      
         Set p = p.Container

      Loop
      
      If bSkip = False Then
         
         TName = LCase(TypeName(Ctrl))
         CName = LCase(Ctrl.Name)
   
         Select Case TName
            Case "textbox":
               Call EnableTxt(Ctrl, bEnable)
            Case "combobox":
               Ctrl.Locked = Not bEnable
            Case "listbox", "checkbox", "optionbutton":
               Ctrl.Enabled = bEnable
            Case "commandbutton":
               'If Ctrl.Cancel = False And CName <> "bt_cancel" And CName <> "bt_close" And CName <> "bt_cerr" Then
               If Ctrl.Cancel = False And InStr(1, CmdEsp, "," & CName & ",", vbTextCompare) = 0 Then
                  Ctrl.Enabled = bEnable
               End If
            Case "spinbutton":
               Ctrl.visible = bEnable
            Case "gred":
               Ctrl.Enabled = bEnable
            Case "fed4grid", "fed3grid", "fed2grid":
               Ctrl.Locked = Not bEnable
            Case "fedgrid":
               Ctrl.Enabled = bEnable
            
            Case "label", "frame", "line", "commondialog":
               ' nada nada
            Case Else:
               If InStr(1, MsgNotSup, TName) <= 0 Then
                  MsgNotSup = MsgNotSup & "," & TName
                  Debug.Print "EnableForm0: control " & TName & " no soportado."
               End If
         End Select
      End If
   Next i
End Sub


Sub EnableForm0(Frm As Form, ByVal bEnable As Boolean)
   Dim i As Integer, CmdEsp As String
   Dim TName As String, CName As String
   Dim Ctrl As Control
   Static MsgNotSup As String

   ' Estos botones no los modifica
   CmdEsp = ",bt_cancel,bt_cancelar,bt_close,bt_cerrar,bt_print,bt_imprimir,bt_preview,bt_copyexcel,"

   For i = 0 To Frm.Controls.Count - 1

      Set Ctrl = Frm.Controls(i)
      TName = LCase(TypeName(Ctrl))
      CName = LCase(Ctrl.Name)

      Select Case TName
         Case "textbox":
            Call EnableTxt(Ctrl, bEnable)
         Case "combobox":
            Ctrl.Locked = Not bEnable
         Case "listbox", "checkbox", "optionbutton":
            Ctrl.Enabled = bEnable
         Case "commandbutton":
            'If Ctrl.Cancel = False And CName <> "bt_cancel" And CName <> "bt_close" And CName <> "bt_cerr" Then
            If Ctrl.Cancel = False And InStr(1, CmdEsp, "," & CName & ",", vbTextCompare) = 0 Then
               Ctrl.Enabled = bEnable
            End If
         Case "spinbutton":
            Ctrl.visible = bEnable
         Case "gred":
            Ctrl.Enabled = bEnable
         Case "fed4grid", "fed3grid", "fed2grid":
            Ctrl.Locked = Not bEnable
         Case "fedgrid":
            Ctrl.Enabled = bEnable
         
         Case "label", "frame", "line", "commondialog":
            ' nada nada
         Case Else:
            If InStr(1, MsgNotSup, TName) <= 0 Then
               MsgNotSup = MsgNotSup & "," & TName
               Debug.Print "EnableForm0: control " & TName & " no soportado."
            End If
      End Select
   Next i

End Sub

Sub EnableTxt(Txt As Control, ByVal bool As Integer)

   Txt.Locked = Not bool

   If bool Then
      Txt.BackColor = vbWhite
   Else
      Txt.BackColor = Txt.Parent.BackColor
   End If

End Sub
' Enum para encontrar App
Private Function EnumWindowsProc2FndApp(ByVal hWind As Long, ByVal lParam As Long) As Integer
   Dim Buf As String * 101, Capt As String
   Dim Rc As Long
   
   Rc = GetWindowText(hWind, Buf, 100)

   If Rc Then
      Capt = UCase(Left(Buf, Rc))
      If InStr(Capt, UCase(lFindApp.Caption)) Then
         lFindApp.hWnd = hWind
         
         EnumWindowsProc2FndApp = 0 ' para no siga

         Exit Function
      End If
   End If
   
   EnumWindowsProc2FndApp = 1 ' para se siga con la siguiente
End Function

'
' Transforma un string dejando solo la primera letra
' en mayusculas
'
Function FCase(ByVal Buf As String, Optional MLen As Byte = 0) As String
   Dim i As Integer, lBuf As Integer, j As Integer, p As Integer
   Dim lCh As String, uCh As String, W As String, lW As Integer

   Buf = Trim(Buf)
   lBuf = Len(Buf)
   If lBuf = 0 Then
      FCase = Buf
      Exit Function
   End If
   
   If gUpCase Then
      FCase = UCase(Buf)
      Exit Function
   End If

   j = -1
   p = 0

   For i = 1 To lBuf
      lCh = LCase(Mid(Buf, i, 1))
      uCh = UCase(lCh)
    
      If (uCh <> "." And uCh = lCh) Or i = lBuf Then  ' no es letra o se termino el buffer
         
         If j > 0 And i - j + 1 > 1 Then
         
            If i = lBuf Then
               W = Mid(Buf, j)
            Else
               W = Mid(Buf, j, i - j)
            End If
            lW = Len(W)
            
            If InStr(1, gLoWords, "," & W & ",", vbTextCompare) > 0 Then
            
               If p > 0 Then
                  W = LCase(W)
               Else
                  W = UCase(Mid(W, 1, 1)) & LCase(Mid(W, 2))
               End If
               
               p = p + 1
            ElseIf (InStr(W, ".") = 0 Or InStr(W, ".") = lW) And InStr(1, gSkWords, "," & W & ",", vbTextCompare) = 0 Then
               W = UCase(Mid(W, 1, 1)) & LCase(Mid(W, 2))
               
               If Right(W, 1) = "." Then
                  p = 0
               Else
                  p = p + 1
               End If
            Else
               W = ""
            End If
                        
            If W <> "" Then
               If i = lBuf Then
                  Mid(Buf, j) = W
               Else
                  Mid(Buf, j, i - j) = W
               End If
            End If
            
         End If
         
         j = -1
      ElseIf j = -1 Then
         j = i
      End If
   Next i

   FCase = Buf

End Function

'
' Transforma un string dejando solo la primera letra
' en mayusculas
'
Function FCase_old(ByVal Buf As String, Optional MLen As Byte = 0) As String
   Dim i As Integer, Ln As Integer, fr As Integer, j As Integer, p As Integer
   Dim lCh As String, uCh As String

   Buf = Trim(Buf)
   Ln = Len(Buf)
   If Ln = 0 Then
      FCase_old = Buf
      Exit Function
   End If
   
   If gUpCase Then
      FCase_old = UCase(Buf)
      Exit Function
   End If

   fr = True

   j = -1
   p = 0

   For i = 1 To Ln
      lCh = LCase(Mid(Buf, i, 1))
      uCh = UCase(lCh)

      If fr Then
            
         Mid(Buf, i, 1) = uCh
         j = i ' inicio de palabra
      Else
         Mid(Buf, i, 1) = lCh
      End If
    
      If uCh = lCh Then ' no es letra
         fr = True
         
         If p > 0 And j > 1 Then
            If InStr(1, gNullWords, "," & Mid(Buf, j, i - j) & ",", vbTextCompare) And Mid(Buf, j - 1, 1) = " " And Mid(Buf, i, 1) = " " Then
            ' If i - j <= MLen And j > 1 Then
               Mid(Buf, j, 1) = LCase(Mid(Buf, j, 1))
            End If
         End If
         
         p = p + 1
      Else
         fr = False
      End If
   Next i

   FCase_old = Buf

End Function

Public Function CbFillAno(cbAno As ComboBox, Optional Sel As Integer = -1, Optional ByVal MinAno As Integer = -1, Optional ByVal MaxAno As Integer = -1)
   Dim Ano As Integer
   
   If MaxAno < 0 Then
      MaxAno = Year(Now) + 1
   ElseIf MaxAno < 100 Then
      MaxAno = Year(Now) + MaxAno
   End If
   
   If Sel > MaxAno Then
      MaxAno = Sel
   ElseIf Sel < MinAno Then
      Sel = MinAno
   End If
   
   If MinAno < 0 Then
      MinAno = MaxAno - 5
   ElseIf MinAno < 100 Then
      MinAno = Year(Now) - MinAno
   End If
         
   For Ano = MaxAno To MinAno Step -1
      Call AddItem(cbAno, Ano, Ano, Ano = Sel)
   Next Ano
   
   If Sel = -1 And cbAno.ListIndex < 0 Then
      cbAno.ListIndex = 0
   End If

End Function

Sub FillMes(CbMes As Control, Optional ByVal SelMes As Integer = -1)
   Debug.Print "*** Cambiar FillMes por CbFillMes ***"
   Call CbFillMes(CbMes, SelMes)
End Sub
Sub CbFillMes(Mes As Control, Optional ByVal SelMes As Integer = -1)
   Dim i As Integer

   For i = 1 To 12
   
      Mes.AddItem gNomMes(i)
      Mes.ItemData(Mes.NewIndex) = i
      
      If i = SelMes Then
         Mes.ListIndex = Mes.NewIndex
      End If
      
   Next i
   
End Sub
' Llena una combo con una lista de valores
Sub FillCombo2(Cb As Control, strArray() As String, Optional ByVal First As Integer = 1, Optional ByVal bFirstDef As Boolean = 1)
   Dim i As Integer
   
   For i = First To UBound(strArray)
      If Len(strArray(i)) > 0 Then
         Call CbAddItem(Cb, strArray(i), i, bFirstDef And (i = First))
      End If
   Next i

End Sub
Private Function FindApp(Cap As String) As Long

   lFindApp.Caption = Cap
   lFindApp.hWnd = 0

   'Call EnumWindowStations(AddressOf EnumWindowsProc2FndApp, 111)
   'Call EnumThreadWindows(0, AddressOf EnumWindowsProc2FndApp, 111)
   Call EnumWindows(AddressOf EnumWindowsProc2FndApp, 111)

   FindApp = lFindApp.hWnd

End Function
' Retorna el 1er y último dia del mes de la fecha
Sub FirstLastMonthDay(ByVal Fecha As Double, First As Long, Last As Long)

   First = DateSerial(Year(Fecha), month(Fecha), 1)
   Last = DateSerial(Year(Fecha), month(Fecha) + 1, 1) - 1

End Sub
Function FmtFecha(ByVal Fecha As Double, Optional ByVal bLong As Byte = 0) As String
   ' "d mmm yy"
   ' "dd-mmm-yy"
   
   If Fecha <= 1000 Then
      FmtFecha = ""
      Exit Function
   End If
   
   If bLong Then
      FmtFecha = Day(Fecha) & IIf(bLong = 2, " ", " de ") & gNomMes(month(Fecha)) & IIf(bLong = 2, " ", " de ") & Year(Fecha)
      'FmtFecha = Format(Fecha, "d \d\e mmmm \d\e yyyy")
   Else
      FmtFecha = Day(Fecha) & " " & Left(gNomMes(month(Fecha)), 3) & " " & Year(Fecha)
      ' FmtFecha = Format(Fecha, "d mmm yyyy")
   End If

End Function
Function FmtAnoMes(ByVal AnoMes As Long, Optional ByVal bLong As Byte = 0) As String
   
   If AnoMes <= 0 Then
      FmtAnoMes = "???"
      Exit Function
   End If

   If bLong Then
      FmtAnoMes = gNomMes(AnoMes Mod 100) & IIf(bLong = 2, " ", " de ") & (AnoMes \ 100)
   Else
      FmtAnoMes = Left(gNomMes(AnoMes Mod 100), 3) & " " & (AnoMes \ 100)
   End If

End Function
Function FmtMes(ByVal Fecha As Double, Optional ByVal bLong As Byte = 0) As String
   ' "mmm yy"
   ' "mmm-yy"
   
   If Fecha <= 0 Then
      FmtMes = "???"
      Exit Function
   End If

   If bLong Then
      FmtMes = gNomMes(month(Fecha)) & IIf(bLong = 2, " ", " de ") & Year(Fecha)
      ' FmtMes = Format(Fecha, "mmmm \d\e yyyy")
   Else
      FmtMes = Left(gNomMes(month(Fecha)), 3) & " " & Year(Fecha)
      ' FmtMes = Format(Fecha, "mmm yyyy")
   End If

End Function
Function FmtDiaSem(ByVal Fecha As Double, Optional ByVal bLong As Boolean = 0) As String
   Dim Nombre As String
   
   Nombre = gDiaSem1(Weekday(Fecha, vbMonday))

   If bLong Then
      FmtDiaSem = Nombre
   Else
      FmtDiaSem = Left(Nombre, 3)
   End If

End Function

Function FmtFecha2_old(ByVal F As Double) As String
   ' "d de mmm de yyyy"
   
   If F Then
      FmtFecha2_old = Day(F) & " de " & gNomMes(month(F)) & " de " & Year(F)
   Else
      FmtFecha2_old = "???"
   End If

End Function


Sub FormColor(Frm As Form)
   Dim i As Integer
   Dim TName As String

   Frm.BackColor = vbButtonFace

   For i = 0 To Frm.Controls.Count - 1

      TName = TypeName(Frm.Controls(i))

      'If TypeOf Frm.Controls(i) Is Label Then
      If TName = "Label" Then
         Frm.Controls(i).BackColor = vbButtonFace
      'ElseIf TypeOf Frm.Controls(i) Is Frame Then
      ElseIf TName = "Frame" Then
         Frm.Controls(i).BackColor = vbButtonFace
      End If
   Next i

End Sub

'
' Para las posiciones de las ventana
'
Sub FormPos(Frm As Form, Optional Align As Integer = -555)

   If Align = -555 Then
      Align = FRMALIGN
   End If

   If Align = vbCenter Then
      Frm.Left = (Screen.Width - Frm.Width) / 2
      Frm.Top = (Screen.Height - Frm.Height) / 2
      
   Else    ' Visible
      If Frm.Top < 0 Or Frm.Top + Frm.Height + W.YCaption > Screen.Height Then
         Frm.Top = (Screen.Height - Frm.Height) - W.YCaption
      End If
   
      If Frm.Left < 0 Or Frm.Left + Frm.Width > Screen.Width Then
         Frm.Left = (Screen.Width - Frm.Width) - W.xFrame
      End If
   
      'If Frm.Top + Frm.Height > Screen.Height - W.YCaption Then
      '   Frm.Top = Screen.Height - Frm.Height - W.YCaption * 2
      'End If
   
      'If Frm.Left + Frm.Width > Screen.Width Then
      '   Frm.Left = Screen.Width - Frm.Width
      'End If
   End If
   
End Sub

' Elimina los items de una colección
Public Sub FreeCol(Col As Collection, Optional ByVal bFree As Boolean = True)

   If Col Is Nothing Then
      Exit Sub
   End If

   Do While Col.Count > 0
      Col.Remove 1
   Loop
   
   If bFree Then
      Set Col = Nothing
   End If

End Sub
' Se sugiere pasar como parámetro Nombre & "#" & Clave. Seed permite hacerlo dependiente del programa
' El numero que retorna es el que se debería guardar en la DB
' Esta funcion es la misma que en PamInc.asp
Public Function GenClave(ByVal Buf As String, Optional ByVal AppSeed As Long = 0) As Long
   Dim i As Long, Ln As Long
   Dim l As Long
   Dim m As Long, c As Long
   
   Ln = Len(Buf)
   l = 134537 + AppSeed
   
   For i = 1 To Ln
      l = l + (Asc(Mid(Buf, i, 1)) * (i + 111)) Mod 9765
   Next i

   m = 87731 + 131 * Ln + 7
   c = l Mod m
   
   GenClave = c
   
End Function
Public Function Hash(ByVal Buf As String) As Long
   Dim i As Integer, lBuf As Integer
   Dim H As Long
   
   lBuf = Len(Buf)
   
   H = 0
   For i = 1 To lBuf
      H = H + Asc(Mid(Buf, i, 1)) * i
   Next i
   
   Hash = H
   
End Function
'
' Trata de encontrar un código unico para cada string
'
Function Hash2(ByVal Buf As String) As Long
   Dim i As Integer
   Dim Sum As Long
   Dim ch As String * 1
   
   Sum = 0
   For i = 1 To Len(Buf)
      ch = Mid(Buf, i, 1)
      If ch = "," Then
      ElseIf IsNumeric(ch) Then
         Sum = Sum + i * Val(ch)
      Else
         Sum = Sum + i * (Asc(ch) - Asc(" "))
      End If
   Next i
   
   Hash2 = Sum

End Function

' Se sugiere pasar como parámetro Nombre & Clave. Seed permite hacerlo dependiente del programa
' El numero que retorna es el que se debería guardar en la DB
Public Function GenClave2(ByVal Buf As String, Optional ByVal AppSeed As Long = 0) As Long
   Dim i As Long, lBuf As Long
   Dim c As Long, d As Double
   
   Debug.Print "*** Conviene usar GenClave3 ****"
   
   lBuf = Len(Buf)
   
   d = 0
   For i = 1 To lBuf
'      Debug.Print i & ") " & Mid(Buf, i, 1) & ": " & Log((Asc(Mid(Buf, i, 1)) * (i + lBuf))) & " = " & d
      d = d + Log((Asc(Mid(Buf, i, 1)) * (i + lBuf)))
   Next i

   d = d + Log(AppSeed + lBuf)

   d = Log(d) * 3759
   c = Int((d - Int(d)) * 3175391)
   c = c + AppSeed / 2
   
   GenClave2 = c
   
End Function

' Se sugiere pasar como parámetro Nombre & Clave. Seed permite hacerlo dependiente del programa
' El numero que retorna es el que se debería guardar en la DB
Public Function GenClave3(ByVal Buf As String, Optional ByVal AppSeed As Long = 0) As Long
   Dim i As Long, lBuf As Long
   Dim c As Long, d As Double
   
   lBuf = Len(Buf)
   
   d = 0
   For i = 1 To lBuf
'      Debug.Print i & ") " & Mid(Buf, i, 1) & ": " & Log((Asc(Mid(Buf, i, 1)) * (i + lBuf))) & " = " & d
      d = d + Log((Asc(Mid(Buf, i, 1)) * (i + lBuf))) * i * 357
   Next i

   d = d * 753 * Log(AppSeed + lBuf)

   d = Log(d) * 3759
   c = Int((d - Int(d)) * 3175391)
   c = c + AppSeed / 2
   
   GenClave3 = c
   
End Function

Function GetMon(Buf As String) As Integer
   Dim i As Integer
   Dim uBuf As String

   uBuf = UCase(Left(Buf, 3))

   For i = 1 To 12
      If UCase(Left(gNomMes(i), 3)) = uBuf Then
         GetMon = i
         Exit Function
      End If

      If UCase(Left(gMonNam(i), 3)) = uBuf Then
         GetMon = i
         Exit Function
      End If
   Next i

   GetMon = 0

End Function

Public Function GetIniString(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal DefValue As String = "") As String
   Dim Aux As String
   Dim Rc As Integer

   Aux = Space(5000)

   Rc = GetPrivateProfileString(Section, Key, DefValue, Aux, Len(Aux) - 1, IniFile)
   GetIniString = Trim(Left(Aux, Rc))

End Function
' Si Value = vbNullString se elimina la Key
' Si Key = vbNullString, se elimina la Section
' Ojo: "" es distinto de vbNullString
Public Function SetIniString(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
   Dim Rc As Long

   Rc = WritePrivateProfileString(Section, Key, Value, IniFile)
   SetIniString = (Rc <> 0)
      
   Rc = FlushProfile(0, 0, 0, IniFile)

End Function

' 2 ago 2019: por si los usuarios ingresan leseras
Private Function MyDateSerial(ByVal Ano As Integer, ByVal Mes As Integer, ByVal Dia As Integer)

   If (Ano < 1901 Or Ano > 9900) Or (Mes < 0 Or Mes > 900) Or (Dia < 0 Or Dia > 900) Then
      MyDateSerial = -1
   Else
      MyDateSerial = DateSerial(Ano, Mes, Dia)
   End If
End Function

'
' Para fecha del tipo    D/M/YYYY o D-M-YYYY o DDMMYY
'
Function GetDate(ByVal Dt As String, Optional ByVal Fmt As String = "dmy", Optional bHour As Boolean = 0) As Double
   Dim s1 As Integer, s2 As Integer, Sep As String * 1
   Dim a As Integer, Hr As String, Tm As Double, LnDt As Integer
   Dim P1 As Integer, p2 As Integer, p3 As Integer

   If Trim(Dt) = "" Then
      GetDate = 0
      Exit Function
   End If

   If Val(Dt) <= 0 Then ' 23 abr 2018
      GetDate = -2
      Exit Function
   End If

   Dt = Trim(Dt)
   LnDt = Len(Dt)
   
   ' por si vienen horas h:m
   s1 = InStr(1, Dt, " ", vbBinaryCompare)
   If s1 Then
      Hr = Trim(Mid(Dt, s1 + 1))
      Dt = Trim(Left(Dt, s1 - 1))
   End If
   
   s1 = InStr(Dt, "/")
   If s1 = 0 Then
      s1 = InStr(Dt, "-")
      If s1 Then
         Sep = "-"
      End If
   Else
      Sep = "/"
   End If
   
   If s1 = 0 Then ' Viene sin separadores, se asumen formatos fijos
      Sep = "/"
      If Len(Fmt) = 3 Then
         If LnDt = 4 Then  ' DDMM
            Dt = Left(Dt, 2) & "/" & Right(Dt, 2)
         ElseIf LnDt = 6 Or LnDt = 8 Then ' DDMMYY or 'DDMMYYYY
            Dt = Left(Dt, 2) & "/" & Mid(Dt, 3, 2) & "/" & Mid(Dt, 5)
         ElseIf LnDt = 1 Or LnDt = 2 Then ' D or DD
            Dt = Dt & "/" & month(Now) & "/" & Year(Now)
         Else
            Sep = ""
         End If
      Else
         If LnDt = 4 Or LnDt = 6 Then  ' MMYY or MMYYYY
            Dt = Left(Dt, 2) & "/" & Mid(Dt, 3)
         ElseIf LnDt = 3 Or LnDt = 5 Then  ' MYY or MYYYY
            Dt = Left(Dt, 1) & "/" & Mid(Dt, 2)
         Else
            Sep = ""
         End If
      End If
   
      s1 = InStr(Dt, Sep)

      If s1 <= 0 Then ' 6 jun 2017: no calza con nada
         GetDate = 0
         Exit Function
      End If
         
   End If
   
   If Len(Fmt) >= 2 And Len(Fmt) <= 3 Then
      If Sep = "" Then
         GetDate = -1
         Exit Function
      End If
      
      s1 = InStr(Dt, Sep)
      If s1 > 0 Then
         P1 = Val(Left(Dt, s1 - 1))
      End If
      
      s2 = InStr(s1 + 1, Dt, Sep)
      If s2 > 0 Then
         p2 = Val(Mid(Dt, s1 + 1, s2 - s1))
         p3 = Val(Mid(Dt, s2 + 1))
      Else
         p2 = Val(Mid(Dt, s1 + 1))
         p3 = 0
      End If
      
      If LCase(Fmt) = "dmy" Then
         If p3 = 0 Then
            p3 = Year(Now)
         ElseIf p3 < 50 Then
            p3 = p3 + 2000
         ElseIf p3 < 100 Then
            p3 = p3 + 1900
         End If
         
         If p2 = 0 Then
            p2 = month(Now)
         End If
               
         Tm = MyDateSerial(p3, p2, P1)
      ElseIf LCase(Fmt) = "mdy" Then
         If p3 = 0 Then
            p3 = Year(Now)
         ElseIf p3 < 50 Then
            p3 = p3 + 2000
         ElseIf p3 < 100 Then
            p3 = p3 + 1900
         End If
                  
         Tm = MyDateSerial(p3, P1, p2)
      ElseIf LCase(Fmt) = "ymd" Then
         Tm = MyDateSerial(P1, p2, p3)
      ElseIf LCase(Fmt) = "my" Then
         If p2 < 50 Then
            p2 = p2 + 2000
         ElseIf p2 < 100 Then
            p2 = p2 + 1900
         End If
      
         Tm = MyDateSerial(p2, P1, 1)
      ElseIf LCase(Fmt) = "dm" Then
         Tm = MyDateSerial(Year(Now), p2, P1)
      End If
   
      If bHour And Tm > 0 Then
         Tm = Tm + GetTime(Hr)
      End If

      GetDate = Tm
      
      Exit Function
   
'      s2 = InStr(s1 + 1, Dt, Sep)
'      If s2 = 0 Then
'         a = Year(Now)
'         s2 = Len(Dt) + 1
'      Else
'         a = ValYear(Mid(Dt, s2 + 1))
'      End If
   Else
      If s1 = 0 Then
         a = Year(Now)
         s1 = Len(Dt) + 1
      Else
         a = ValYear(Mid(Dt, s1 + 1))
      End If
   
   End If
            
   On Error Resume Next
   
   If LCase(Fmt) = "dmy" Then
      Tm = MyDateSerial(a, Val(Mid(Dt, s1 + 1, s2 - s1 - 1)), Val(Left(Dt, s1 - 1)))
   ElseIf LCase(Fmt) = "mdy" Then
      Tm = MyDateSerial(a, Val(Left(Dt, s1 - 1)), Val(Mid(Dt, s1 + 1, s2 - s1 - 1)))
   ElseIf LCase(Fmt) = "ymd" Then
      Tm = MyDateSerial(a, Val(Left(Dt, s1 - 1)), Val(Mid(Dt, s1 + 1, s2 - s1 - 1)))
   ElseIf LCase(Fmt) = "my" Then
      Tm = MyDateSerial(a, Val(Left(Dt, s1 - 1)), 1)
   End If
   
   If bHour And Tm > 0 Then
      Tm = Tm + GetTime(Hr)
   End If
   
   GetDate = Tm
   
   If ERR Then
      GetDate = -1
   End If

End Function



' Fmt debe ir algo del tipo  yyyy-mm-dd o yy-mm-dd o yyyy-mm-dd hh:nn pp
'
Public Function GetFixDate(ByVal Fld As String, ByVal Fmt As String, ByVal Def As String, DefUsed As Integer) As Double
   Dim i As Integer, j As Integer, l As Integer, lf As Integer
   Dim F(6) As String, d(6) As Integer, Dt As Double
   
   If Trim(Fld) = "" Or Trim(Fmt) = "" Then
      GetFixDate = 0
      Exit Function
   End If
   
   Fmt = LCase(Fmt) & " "
   F(0) = "y"
   F(1) = "m"
   F(2) = "d"
   F(3) = "h"
   F(4) = "n"
   F(5) = "s"
   F(6) = "p"

   lf = Len(Fmt)
   DefUsed = 0

   For i = 0 To UBound(F)
   
      j = InStr(Fmt, F(i))

      If j Then
         For l = j + 1 To lf
            If Mid(Fmt, l, 1) <> F(i) Then
            
               If i <> 6 Then  ' pp
                  d(i) = Val(Mid(Fld, j, l - j))
               ElseIf StrComp(Mid(Fld, j, l - j), "pm", vbTextCompare) = 0 Then
                  d(6) = 1 ' PM
               End If
               
               If i <= 2 And d(i) <= 0 Then
                  d(i) = Val(Mid(Def, j, l - j))
                  DefUsed = 1
               ElseIf d(i) < 0 Then
                  d(i) = Val(Mid(Def, j, l - j))
                  DefUsed = 1
               End If
               
               If i <= 2 And d(i) <= 0 Then
                  GetFixDate = -1
                  Exit Function
               ElseIf d(i) < 0 Then
                  GetFixDate = -1
                  Exit Function
               End If
   
               Exit For
            End If
         Next l
      End If
   Next i

   If d(0) <= 0 Or d(1) <= 0 Or d(2) <= 0 Then
      GetFixDate = -1
      Exit Function
   End If

   If d(0) < 100 Then
      If d(0) < 80 Then
         d(0) = d(0) + 2000
      Else
         d(0) = d(0) + 1900
      End If
   End If
   
   Dt = DateSerial(d(0), d(1), d(2))
   
   If d(3) > 0 Or d(4) > 0 Or d(5) > 0 Then
   
      If d(6) Then
         d(3) = d(3) + 12
      End If
   
      Dt = Dt + TimeSerial(d(3), d(4), d(5))
   End If

   GetFixDate = Dt
   
End Function
'
' Para horas del tipo:  hh:nn[:ss][PM] o h:nn[:ss][PM]
'
Function GetTime(ByVal Hr As String) As Double
   Dim i As Integer, j As Integer, k As Integer
   Dim H As Integer, m As Integer, s As Integer

   If Trim(Hr) = "" Then
      GetTime = 0
      Exit Function
   End If

   i = InStr(1, Hr, ":", vbBinaryCompare)
   H = Val(Left(Hr, i - 1))
   
   j = InStr(i + 1, Hr, ":", vbBinaryCompare)
   If j Then
      m = Val(Mid(Hr, i + 1, j - i - 1))
      s = Val(Mid(Hr, j + 1, 2))
   Else
      m = Val(Mid(Hr, i + 1))
      s = 0
   End If
   
   If H <> 12 And InStr(i + 1, Hr, "PM", vbTextCompare) Then
      H = H + 12
   End If
   
   GetTime = TimeSerial(H, m, s)

End Function
' Filtra el ingreso de caracteres para codigos (identificadores)
' No acepta espacios o caracteres raros
' Debería usarse en el evento KeyPress
Sub KeyCod(KeyAscii As Integer)
   Dim ch As String * 1
   Dim Codes As String
   
   ch = LCase(Chr(KeyAscii))

   'Codes = "0123456789ñüáéíóú-_"
   Codes = "0123456789-_"

   'If KeySys(KeyAscii) = False And Not IsNumeric(Ch) And (Ch < "a" Or Ch > "z") And Ch <> "-" And Ch <> "_" Then
   If KeySys(KeyAscii) = False And (ch < "a" Or ch > "z") And InStr(Codes, ch) = 0 Then
      Beep
      KeyAscii = 0
   End If
   
End Sub
' Filtra el ingreso de caracteres para emails
' No acepta espacios o caracteres raros
' Debería usarse en el evento KeyPress
Sub KeyMail(KeyAscii As Integer)
   Dim ch As String * 1
   Dim Codes As String
   
   ch = LCase(Chr(KeyAscii))

   Codes = "0123456789@.¡!#$%&’*+-/=¿?^_`{|}~ñüáéíóú"

   'If KeySys(KeyAscii) = False And Not IsNumeric(Ch) And (Ch < "a" Or Ch > "z") And Ch <> "-" And Ch <> "_" Then
   If KeySys(KeyAscii) = False And (ch < "a" Or ch > "z") And InStr(Codes, ch) = 0 And ch <> ";" Then ' 14 ene 2020: se agrega el ; por si hay más de uno
      Beep
      KeyAscii = 0
   End If
   
End Sub
' Filtra el ingreso de caracteres para codigos (identificadores)
' No acepta espacios o caracteres raros
' Debería usarse en el evento KeyPress
' Trsnsforma a mayúsculas
Sub KeyUCod(KeyAscii As Integer)

   Call KeyCod(KeyAscii)
   If KeyAscii Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If

End Sub
' Permite saber si el usuario precionó un COPY (Ctrl+C o Ctrl+Ins), retorna True o False
' Debería usarse en el evento KeyDown
' Tipicamente se usa en la grilla
Function KeyCopy(ByVal KeyCode As Integer, ByVal Shift As Integer) As Integer
   KeyCopy = (Shift = vbCtrlMask And (UCase(Chr(KeyCode)) = "C" Or KeyCode = vbKeyInsert))
End Function
' Permite saber si el usuario precionó un Ctrl+<letra>
' Debería usarse en el evento KeyDown
Function KeyCtrl(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal Letra As String) As Integer
   KeyCtrl = (Shift = vbCtrlMask And UCase(Chr(KeyCode)) = UCase(Letra))
End Function
' Filtra el ingreso de caracteres para fechas (numeros y '/')
' Debería usarse en el evento KeyPress
Sub KeyDate(KeyAscii As Integer)
   Dim ch As String * 1

   ch = Chr(KeyAscii)
   
   If KeySys(KeyAscii) = False And Not IsNumeric(ch) And ch <> "/" Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Filtra el ingreso de caracteres para numeros con decimales (numeros, '.' y ',')
' Debería usarse en el evento KeyPress
Sub KeyDec(KeyAscii As Integer)
   Dim ch As String * 1
   
   ch = Chr(KeyAscii)

   If KeyAscii = vbKeyReturn Or (KeySys(KeyAscii) = False And Not IsNumeric(ch) And ch <> lSigDec And ch <> "-") Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Filtra el ingreso de caracteres para numeros positivos con decimales (numeros, '.' y ',')
' Debería usarse en el evento KeyPress
Sub KeyDecPos(KeyAscii As Integer)
   Dim ch As String * 1
   
   ch = Chr(KeyAscii)

   If KeyAscii = vbKeyReturn Or (KeySys(KeyAscii) = False And Not IsNumeric(ch) And ch <> lSigDec) Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Filtra el ingreso de caracteres para numeros con decimales (numeros, '.' y ',' y %)
' Debería usarse en el evento KeyPress
Sub KeyPorc(KeyAscii As Integer)
   Dim ch As String * 1
   
   ch = Chr(KeyAscii)

   If KeyAscii = vbKeyReturn Or (KeySys(KeyAscii) = False And Not IsNumeric(ch) And ch <> lSigDec And ch <> "-" And ch <> "%") Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Filtra el ingreso de caracteres para numeros hexadecimales (numeros y letras de la A a la F)
' Debería usarse en el evento KeyPress
Sub KeyHex(KeyAscii As Integer)
   Dim ch As String * 1

   ch = UCase(Chr(KeyAscii))

   If KeySys(KeyAscii) = False And Not IsNumeric(ch) And (ch < "A" Or ch > "F") Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Transforma todo lo ingresado a minúsculas
' Debería usarse en el evento KeyPress
Sub KeyLower(KeyAscii As Integer)

   KeyAscii = Asc(LCase(Chr(KeyAscii)))

End Sub
' Filtra el ingreso de caracteres para textos, pero sólo caracteres imprimibles
' Debería usarse en el evento KeyPress
Public Sub KeyName(KeyAscii As Integer)
   Dim ch As String * 1
   Dim i As Integer
   
   ch = LCase(Chr(KeyAscii))

   If KeySys(KeyAscii) = False And (KeyAscii < 32 Or KeyAscii > 127) Then
      
      If KeyAscii > 127 Then
         For i = 0 To 5
            If Asc(ChrMask(i, 0)) = KeyAscii Then
               Beep
               KeyAscii = 0
               Exit Sub
            End If
         Next i
         
      Else
         Beep
         KeyAscii = 0
      End If
      
   End If

   'Dim i As Integer
   'For i = 1 To 5
   '   If KeyAscii = Asc(ChrMask(i, 0)) Then
   '      MsgBeep vbExclamation
   '      KeyAscii = 0
   '      Exit Sub
   '   End If
   'Next i

End Sub
Function KeyMenu(KeyCode As Integer, Shift As Integer)

   KeyMenu = (KeyCode = KEY_MENU95)
   
End Function

' Revisa si es una tecla del sistema que DEBE dejar pasar
Function KeySys(ByVal KeyAscii As Integer) As Boolean

   ' 3: Copy
   ' 22: Paste
   ' 24: Cut

   '              Back                  Copy           Paste             Cut            Undo (ctrl-Z)
   KeySys = (KeyAscii = vbKeyBack Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Or KeyAscii = 26)

End Function

' Filtra el ingreso de caracteres para numeros (no decimales), positivos y negativos
' Debería usarse en el evento KeyPress
Sub KeyNum(KeyAscii As Integer)

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   ' If KeyAscii = vbKeyReturn Or (KeySys(KeyAscii) = False And Not IsNumeric(Chr(KeyAscii)) And Not Chr(KeyAscii) = "-") Then
   If KeyAscii = vbKeyReturn Or (Not IsNumeric(Chr(KeyAscii)) And Not Chr(KeyAscii) = "-") Then
      Beep
      KeyAscii = 0
   End If

End Sub
Sub KeyNumDecimal(KeyAscii As Integer)

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   ' If KeyAscii = vbKeyReturn Or (KeySys(KeyAscii) = False And Not IsNumeric(Chr(KeyAscii)) And Not Chr(KeyAscii) = "-") Then
   If KeyAscii = vbKeyReturn Or (Not IsNumeric(Chr(KeyAscii)) And Not Chr(KeyAscii) = ",") Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Ingreso de numeros positivos
Sub KeyNumPos(KeyAscii As Integer)

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   If KeyAscii = vbKeyReturn Or Not IsNumeric(Chr(KeyAscii)) Then
      Beep
      KeyAscii = 0
   End If

End Sub
Sub NumKeyPress(KeyAscii As Integer)
   
   Call KeyNum(KeyAscii)

   'If KeySys(KeyAscii) = False And Not IsNumeric(Chr(KeyAscii)) Then
   '   beep
   '   KeyAscii = 0
   'End If

End Sub

' Permite saber si el usuario precionó un PASTE (Ctrl+V o Shift+Ins), retorna True o False
' Debería usarse en el evento KeyDown
Function KeyPaste(ByVal KeyCode As Integer, ByVal Shift As Integer) As Integer
   KeyPaste = (Shift = vbCtrlMask And UCase(Chr(KeyCode)) = "V") Or (Shift = vbShiftMask And KeyCode = vbKeyInsert)
End Function
' Filtra el ingreso de caracteres para RUTs (números, '.', '-' y 'K')
' Debería usarse en el evento KeyPress
Sub KeyCID(KeyAscii As Integer)

   If gValidRut Then
      Call KeyRut(KeyAscii)
   End If

End Sub
' Filtra el ingreso de caracteres para RUTs (números, '.', '-' y 'K')
' Debería usarse en el evento KeyPress
Sub KeyRut(KeyAscii As Integer)
   Dim ch As String * 1

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   ch = UCase(Chr(KeyAscii))

   If Not IsNumeric(ch) And ch <> "K" And ch <> "-" And ch <> "." Then
      Beep
      KeyAscii = 0
   Else
      KeyAscii = Asc(ch)
   End If

End Sub
' Ingreso de horas (hh:nn)
Sub KeyHour(KeyAscii As Integer)
   Dim ch As String

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   ch = Chr(KeyAscii)
   If KeyAscii = vbKeyReturn Or Not (IsNumeric(ch) Or ch = ":") Then
      Beep
      KeyAscii = 0
   End If

End Sub
' Ingreso de teléfonos
Sub KeyTel(KeyAscii As Integer)
   Dim ch As String

   Const Chars = " (),+-"

   If KeySys(KeyAscii) Then
      Exit Sub
   End If

   ch = Chr(KeyAscii)
   If KeyAscii = vbKeyReturn Or Not (IsNumeric(ch) Or InStr(Chars, ch) > 0) Then
      Beep
      KeyAscii = 0
   End If

End Sub

' Transforma todo lo ingresado a mayúsculas
' Debería usarse en el evento KeyPress
Sub KeyUpper(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
Sub KeyAlpha(KeyAscii As Integer)
   Dim ch As String
   
   If KeySys(KeyAscii) Then
      Exit Sub
   End If
   
   ch = Chr(KeyAscii)
   If UCase(ch) = LCase(ch) Then
      KeyAscii = 0
   End If
   
End Sub
'
' Elimina los & y : del texto
'
Function NoAmp(ByVal Buf As String) As String
   Dim i As Integer

   i = InStr(Buf, "&")
   If i > 0 Then
      Buf = Left(Buf, i - 1) & Mid(Buf, i + 1)
   End If

   i = InStr(Buf, ":")
   If i > 0 Then
      NoAmp = Left(Buf, i - 1) & Mid(Buf, i + 1)
   Else
      NoAmp = Buf
   End If

End Function

Function NotNullStr(ByVal Value As String) As String
   
   If Trim(Value) = "" Then
      NotNullStr = " "
   Else
      NotNullStr = Trim(Value)
   End If

End Function

' Esta función la llama directamente la Num2Words
Private Sub InitNum2Words(Num() As String, Dec() As String)

   Num(0) = " cero"
   Num(1) = " uno"
   Num(2) = " dos"
   Num(3) = " tres"
   Num(4) = " cuatro"
   Num(5) = " cinco"
   Num(6) = " seis"
   Num(7) = " siete"
   Num(8) = " ocho"
   Num(9) = " nueve"

   Dec(0) = ""
   Dec(1) = " diez"
   Dec(2) = " veinte"
   Dec(3) = " treinta"
   Dec(4) = " cuarenta"
   Dec(5) = " cincuenta"
   Dec(6) = " sesenta"
   Dec(7) = " setenta"
   Dec(8) = " ochenta"
   Dec(9) = " noventa"

End Sub


Public Function Num2Words(ByVal l As Double, Optional ByVal bMoneda As Boolean = 0, Optional ByVal nDec As Byte = 0) As String
   Static Num(10) As String, Dec(10) As String
   Dim s As String, s1 As String
   Dim i As Integer, d As Double

   If Num(0) = "" Then
      Call InitNum2Words(Num, Dec)
   End If

   If l = 0 Then
      Num2Words = "cero"
      Exit Function
   End If

   d = Frac(l)
   If d > 0 Then
      If nDec > 0 Then
         s = Round(d, nDec)
      Else
         s = d
      End If
      
      d = Round(d * 10 ^ (Len(s) - 2))
      
      Num2Words = Num2Words(Int(l), bMoneda) & " coma " & Num2Words(d, bMoneda)
      Exit Function
   End If

   s = ""

   Do While l > 0

      If l >= 2000000 Then
         s1 = Num2Words(Int(l / 1000000))
         If Right(s1, 3) = "uno" Then
            s1 = Left(s1, Len(s1) - 1)
         End If
         s = s & " " & s1 & " millones"
         l = l - Int(l / 1000000) * 1000000
      ElseIf l >= 1000000 Then
         s = s & "un millón"
         l = l Mod 1000000
      ElseIf l >= 2000 Then
         s1 = Num2Words(Int(l / 1000))
         If Right(s1, 3) = "uno" Then
            s1 = Left(s1, Len(s1) - 1)
         End If
         s = s & " " & s1 & " mil"
         l = l Mod 1000
      ElseIf l >= 1000 Then
         s = s & " mil"
         l = l Mod 1000
      ElseIf l = 100 Then
         s = s & " cien"
         l = 0
      ElseIf Int(l / 100) = 5 Then
         s = s & " quinientos"
         l = l Mod 100
      ElseIf Int(l / 100) = 7 Then
         s = s & " setecientos"
         l = l Mod 100
      ElseIf Int(l / 100) = 9 Then
         s = s & " novecientos"
         l = l Mod 100
      ElseIf l >= 200 Then
         s = s & Num(Int(l / 100)) & "cientos"
         l = l Mod 100
      ElseIf l > 100 Then
         s = s & " ciento"
         l = l Mod 100
      ElseIf l >= 30 Then
         s = s & Dec(Int(l / 10))
         l = l Mod 10
         
         If l > 0 Then
            s = s & " y"
         End If

      ElseIf l > 20 Then
         s = s & " veinti" & Trim(Num(l - 20))
         l = 0
      ElseIf l = 20 Then
         s = s & Dec(Int(l / 10))
         l = l Mod 10
      ElseIf l = 10 Then
         s = s & " diez"
         l = 0
      ElseIf l = 11 Then
         s = s & " once"
         l = 0
      ElseIf l = 12 Then
         s = s & " doce"
         l = 0
      ElseIf l = 13 Then
         s = s & " trece"
         l = 0
      ElseIf l = 14 Then
         s = s & " catorce"
         l = 0
      ElseIf l = 15 Then
         s = s & " quince"
         l = 0
      ElseIf l > 15 Then
         s = s & " dieci" & Trim(Num(l - 10))
         l = 0
      ElseIf l >= 10 Then
         s = s & " dieci"
         l = 0
      Else
         s = s & Num(l)
         
         l = 0
      End If
   Loop

   If bMoneda And Right(s, 3) = "uno" Then
      s = Left(s, Len(s) - 1)
   End If

   Num2Words = Trim(s)
   
End Function

Sub PamInit()
   Dim l As Long, Rc As Long, i As Integer, j As Integer
   Dim Buff As String * 101, Buf As String, AppName As String
   Dim sDefDec As String, sDefThous As String

   If gFrmMain Is Nothing Then
      Debug.Print "*** Falta asignar el gFrmMain ****"
   End If
   
   W.FStart = Now

   Call RegUnhideFileExt

   ' String básicos
   CR = Chr(13)
   NL = Chr(10)
   Tb = Chr(9)
   CRNL = Chr(13) & Chr(10)

   lSigDec = Mid(Format(1.1, "0.0"), 2, 1) ' para saber que signo se usa como decimal
   'lSigDec = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal")
   
   ' Dimensiones en las ventanas
   
   W.YCaption = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
   W.yMenu = GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY
   W.xFrame = GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
   W.yFrame = GetSystemMetrics(SM_CYFRAME) * Screen.TwipsPerPixelY
   W.xScroll = (GetSystemMetrics(SM_CXVSCROLL) + 2) * Screen.TwipsPerPixelX
   W.yScroll = (GetSystemMetrics(SM_CYHSCROLL) + 2) * Screen.TwipsPerPixelY
   W.WinDir = Environ("windir")
   W.WinDrv = Left(W.WinDir, 2)
   
   l = GetDialogBaseUnits()
   W.dx = (l And &HFFFF&) * Screen.TwipsPerPixelX / 4
   W.dy = ((l And &HFFFF0000) / &HFFFF&) * Screen.TwipsPerPixelY / 8

   If LCase(Format(DateSerial(2000, 1, 1), "mmm")) = "jan" Then ' inglés ?
      sDefDec = "."
      sDefThous = ","
   Else
      sDefDec = ","
      sDefThous = "."
   End If

   W.NumDecSym = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal", sDefDec)
   W.CurDecSym = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep", sDefDec)
   W.NumSepSym = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sThousand", sDefThous)
   W.CurSepSym = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonThousandSep", sDefThous)

   W.ShortDate = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate", "M/d/yyyy")
   W.DateSep = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sDate", "/")

   Buf = QryRegValue(HKEY_CLASSES_ROOT, "mailto\shell\open\command", "", "")
   If Buf <> "" Then
      i = InStr(1, Buf, ".exe", vbTextCompare)
      j = rInStr(Buf, "\", i, vbTextCompare)
      W.DefMailer = LCase(Mid(Buf, j + 1, i + 3 - j))
      
      If W.DefMailer = "msimn.exe" Then
         W.DefMailer = "outlookexpress"
      End If
      
   End If

   gGrayText = GetSysColor(17)

   ' Directorio de la aplicación
   If Right(App.Path, 1) = "\" Then    ' C:\
      gAppPath = Left(App.Path, Len(App.Path) - 1)
   Else
      gAppPath = App.Path
   End If
   'AppPath = gAppPath
   W.AppPath = gAppPath

   If GetFullPath(W.AppPath, Buf) = 0 Then
      W.AppPath = Buf
   End If

   ' Datos del usuario
   W.UserName = GetUserName()
   W.PcName = GetComputerName()
   W.Mac = GetMac()

   ' Directorio temporal
   gTmpDir = Environ("TEMP")
   If gTmpDir = "" Then
      gTmpDir = Environ("TMP")
      If gTmpDir = "" Then
         gTmpDir = "C:"
      End If
   End If
   W.TmpDir = gTmpDir
   
   W.DownDir = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}")
   If W.DownDir = "" Then
      W.DownDir = Environ("USERPROFILE") & "\Downloads"
   End If
   
   On Error Resume Next ' 4 sep 2017: desde Win 7 puede caerse al tratar de obtener la versión
   W.Version = App.Major & "." & App.Minor & "." & App.Revision
   Buf = Right(Trim(App.ProductName), 8)
   
   If Len(Buf) = 8 And IsNumeric(Buf) = True Then  'yyyymmdd
      W.FVersion = DateSerial(Val(Left(Buf, 4)), Val(Mid(Buf, 5, 2)), Val(Right(Buf, 2)))
      AppName = Trim(Left(App.ProductName, Len(App.ProductName) - 8))
   ElseIf Len(Buf) = 6 And IsNumeric(Buf) = True Then 'yymmdd - 10 ene 2020
      i = Val(Left(Buf, 2))
      If i < 90 Then
         i = 2000 + i
      Else
         i = 1900 + i
      End If
      
      W.FVersion = DateSerial(i, Val(Mid(Buf, 3, 2)), Val(Right(Buf, 2)))
      AppName = Trim(Left(App.ProductName, Len(App.ProductName) - 6))

   Else
      AppName = App.ProductName
      W.FVersion = 0
   End If
   
   If W.FVersion < Int(Now) - 10 Or W.FVersion > Int(Now) + 10 Then
      Debug.Print "*** OJO *** Version: " & W.Version & " - " & Format(W.FVersion, "d mmm yyyy") & " ***"
   End If
   
   Call AddLog(App.EXEName & ".exe - " & AppName & " v." & W.Version & " - " & Format(W.FVersion, "d mmm yyyy") & ", PC: " & W.PcName & ", User: " & W.UserName & ", pid:" & GetCurrentProcessId())

   ' Nombres de Meses
   gNomMes(1) = "Enero"
   gNomMes(2) = "Febrero"
   gNomMes(3) = "Marzo"
   gNomMes(4) = "Abril"
   gNomMes(5) = "Mayo"
   gNomMes(6) = "Junio"
   gNomMes(7) = "Julio"
   gNomMes(8) = "Agosto"
   gNomMes(9) = "Septiembre"
   gNomMes(10) = "Octubre"
   gNomMes(11) = "Noviembre"
   gNomMes(12) = "Diciembre"

   gMonNam(1) = "January"
   gMonNam(2) = "February"
   gMonNam(3) = "March"
   gMonNam(4) = "April"
   gMonNam(5) = "May"
   gMonNam(6) = "June"
   gMonNam(7) = "July"
   gMonNam(8) = "August"
   gMonNam(9) = "September"
   gMonNam(10) = "October"
   gMonNam(11) = "November"
   gMonNam(12) = "December"

   ' Nombres de Días
   gDiaSem(1) = "Domingo"
   gDiaSem(2) = "Lunes"
   gDiaSem(3) = "Martes"
   gDiaSem(4) = "Miércoles"
   gDiaSem(5) = "Jueves"
   gDiaSem(6) = "Viernes"
   gDiaSem(7) = "Sábado"

   ' Para Weekday partiendo con el vbMonday
   gDiaSem1(1) = "Lunes"
   gDiaSem1(2) = "Martes"
   gDiaSem1(3) = "Miércoles"
   gDiaSem1(4) = "Jueves"
   gDiaSem1(5) = "Viernes"
   gDiaSem1(6) = "Sábado"
   gDiaSem1(7) = "Domingo"

   gSiNo(VAL_NO) = "No"
   gSiNo(VAL_SI) = "Si"
   gSiNo(VAL_OTRO) = ""

   ' Caracteres para enmascarar los SQL
   ChrMask(0, 0) = "|":      ChrMask(0, 1) = Chr(255)
   ChrMask(1, 0) = Chr(10): ChrMask(1, 1) = Chr(188)
   ChrMask(2, 0) = Chr(13): ChrMask(2, 1) = Chr(189)
   ChrMask(3, 0) = Chr(9):  ChrMask(3, 1) = Chr(163)
   ChrMask(4, 0) = "'":      ChrMask(4, 1) = Chr(164)
   ChrMask(5, 0) = """":     ChrMask(5, 1) = Chr(165)

   'Buf = ReplaceStr(Buf, Chr(253), "'")        ' reemplaza 253 por '
   'Buf = ReplaceStr(Buf, Chr(254), Chr(34))   ' reemplaza 252 por "
   'Buf = ReplaceStr(Buf, Chr(255), "|")        ' reemplaza 251 por |
   'Buf = ReplaceStr(Buf, Chr(248), Chr(10))   ' reemplaza 250 por nl
   'Buf = ReplaceStr(Buf, Chr(247), Chr(13))   ' reemplaza 249 por cr

   ' Para FCase
   gLoWords = ",el,la,los,las,de,del,y,a,o,e,u,en,por,para,"   ' siempre lcase
   gSkWords = ",i,ii,iii,iv,v,x,sa,iva,rut," ' no se consideran
   gNullWords = gLoWords
   
    
   ERR.Clear
   Debug.Print 1 / 0
    
   W.InDesign = (ERR <> 0)
   ERR.Clear

   Set gPrinter = Printer

End Sub

'
' Tranforma una Fecha INFORMIX a una fecha VB
'
' Obtiene el serial de una fecha INFORMIX
' 'yyyy-mm'
'
Function PerVal(Fecha As String) As Long

   If Len(Fecha) <> 7 Then
      PerVal = 0
      Exit Function
   End If

   PerVal = DateSerial(Val(Mid(Fecha, 1, 4)), Val(Mid(Fecha, 6, 2)), 1)

End Function

' retorna cuantos caracteres pudo imprimir en el espacio
Function PrtLine(ByVal Alng As Integer, ByVal lPos As Long, ByVal rPos As Long, ByVal Buf As String, Optional ByVal bTrunc As Boolean = 1, Optional Prt As Object = Nothing) As Integer
   Dim l As Long, r As Long, i As Integer, W As Long, ch As String, j As Integer
   Dim s As Integer
   
   If Prt Is Nothing Then
      Set Prt = Printer
   End If
      
   If lPos <> -1 Then
      l = lPos
   Else
      l = 0
   End If

   If rPos <> -1 Then
      r = rPos
   Else
      r = Prt.Width - 900
   End If

   s = 0
   i = InStr(Buf, vbLf)
   If i Then
      s = 1
   
      If i > 1 Then
         If Mid(Buf, i - 1, 1) = vbCr Then
            s = s + 1
         End If
      End If
      
      Buf = Left(Buf, i - s)
   End If

   W = Prt.TextWidth(Buf)
   If W > r - l Then
   
      For i = 1 To Len(Buf)
                     
         If Prt.TextWidth(Left(Buf, i)) > r - l Then
         
            If bTrunc Then
               Buf = Left(Buf, i - 1)
               PrtLine = i - 1
            Else
               PrtLine = i - 1
               For j = i - 1 To 1 Step -1
                  ch = Mid(Buf, j, 1)
                  If InStr(1, " .,", ch, vbBinaryCompare) Then
                     Buf = Left(Buf, j)
                     PrtLine = j
                     Exit For
                  End If
               
               Next j
            
            End If
            
            Exit For
         End If
         
      Next i
   
      W = Prt.TextWidth(Buf)
   
   Else
      PrtLine = Len(Buf) + s
   End If
   
   If Alng = vbRightJustify Then
      Prt.CurrentX = r - W
   ElseIf Alng = vbCenter Then
      Prt.CurrentX = l + (r - l - W) / 2
   Else ' vbLeft
      Prt.CurrentX = l
   End If

   Prt.Print Buf;
   
End Function

' retorna cuantas lineas pudo imprimir en el espacio
Function PrtBuf(ByVal Alng As Integer, ByVal lPos As Long, ByVal rPos As Long, ByVal Buf As String, Optional Prt As Object = Nothing, Optional bPos As Long = -1) As Integer
   Dim l As Integer, c As Integer
   
   If Prt Is Nothing Then
      Set Prt = Printer
   End If
   
   l = 0
   Do Until Buf = ""
      c = PrtLine(Alng, lPos, rPos, Buf, False, Prt)
      Buf = LTrim(Mid(Buf, c + 1))
      l = l + 1
      If Buf <> "" Then
         Prt.Print
      End If
      
      If bPos > 0 And Prt.CurrentY > bPos Then
         Exit Do
      End If
   Loop
   
   PrtBuf = l
   
End Function

'
' Esta rutina reemplaza en Q1 todas las apariciones de s1 por s2
'
Function ReplaceStr(ByVal Where As String, ByVal FromStr As String, ByVal ToStr As String, Optional ByVal compare As VbCompareMethod = vbTextCompare) As String
   Dim i As Long, j As Long
   Dim L1 As Integer, L2 As Integer

   ' *** falla con un th
   'ReplaceStr = Replace(Where, FromStr, ToStr, , , vbTextCompare)

   L1 = Len(FromStr)
   If L1 <= 0 Then
      ReplaceStr = Where
      Exit Function
   End If
      
   L2 = Len(ToStr)

   i = 1
   Do
      j = InStr(i, Where, FromStr, compare)

      If j = 0 Then
         ReplaceStr = Where
         Exit Function
      End If

      ' Por pato de la InStr con el caracter Chr(253) y la "th"
      If StrComp(Mid(Where, j, L1), FromStr, compare) = 0 Then
         Where = Left(Where, j - 1) & ToStr & Mid(Where, j + L1)
      End If

      i = j + L2

   Loop

End Function

Function FindItem(Cb As Control, ByVal ItemData As Long) As Integer
   FindItem = CbFindItem(Cb, ItemData)
End Function
Function CbFindItem(Cb As Control, ByVal ItemData As Long) As Integer
   Dim i As Integer

   For i = 0 To Cb.ListCount - 1
      If Cb.ItemData(i) = ItemData Then
         CbFindItem = i
         Exit Function
      End If
   Next i

   CbFindItem = -1

End Function

Function SelItem(Cb As Control, ByVal ItemData As Long) As Integer
   SelItem = CbSelItem(Cb, ItemData)
End Function
Function CbSelItem(Cb As Control, ByVal ItemData As Long) As Integer
   Dim i As Long

   i = CbFindItem(Cb, ItemData)
   Cb.ListIndex = i
   
   If TypeName(Cb) = "ListBox" Then
      If i > 0 Then
         Cb.TopIndex = i - 1
      End If
   End If
   CbSelItem = i
   
End Function
Public Function cbMarkSelected(ls As ListBox, ByVal ItemData As Long, ByVal YesNo As Boolean) As Boolean
   Dim i As Integer
   
   i = CbFindItem(ls, ItemData)
   If i > 0 Then
      ls.Selected(i) = YesNo
      cbMarkSelected = True
   End If

End Function

Public Function FindCbText(Cb As Control, ByVal ItemText As String) As Long
   FindCbText = CbFindText(Cb, ItemText)
End Function

Public Function CbFindText(Cb As Control, ByVal ItemText As String) As Long
   Dim i As Integer
   
   For i = 0 To Cb.ListCount - 1
      If StrComp(Cb.list(i), ItemText, vbTextCompare) = 0 Then
         CbFindText = i
         Exit Function
      End If
   Next i
      
   CbFindText = -1
      
End Function

Public Function CbSelText(Cb As Control, ByVal ItemText As String) As Integer
   Dim i As Integer

   i = CbFindText(Cb, ItemText)
   Cb.ListIndex = i
   CbSelText = i
   
End Function

Public Function AddItem(Cb As Control, ByVal ItemText As String, ByVal ItemData As Long, Optional bSelected As Boolean = 0) As Long
   AddItem = CbAddItem(Cb, ItemText, ItemData, bSelected)
End Function
Public Function CbAddItem(Cb As Control, ByVal ItemText As String, ByVal ItemData As Long, Optional ByVal bSelected As Boolean = 0) As Integer
   Dim i As Integer

   Cb.AddItem ItemText
   i = Cb.NewIndex
   Cb.ItemData(i) = ItemData
   CbAddItem = i
   
   If bSelected Then
      Cb.ListIndex = i
   End If
   
End Function
Function ItemData(Cb As Control) As Long
   ItemData = CbItemData(Cb)
   Debug.Print "*** cambiar Itemdata por CbItemData ***"
End Function
Function CbItemData(Cb As Control) As Long

   If Cb.ListIndex >= 0 Then
      CbItemData = Cb.ItemData(Cb.ListIndex)
   Else
      CbItemData = -1
   End If

End Function

Function CbItemDataByte(Cb As Control) As Long
   Dim Data As Long

   Data = CbItemData(Cb)
   If Data < 0 Then
      CbItemDataByte = 0
   Else
      CbItemDataByte = Data
   End If

End Function

Function ItemText(Cb As Control, ByVal ItemData As Long) As String
   ItemText = cbItemText(Cb, ItemData)
End Function
Function cbItemText(Cb As Control, ByVal ItemData As Long) As String
   Dim i As Integer

   i = CbFindItem(Cb, ItemData)
   If i >= 0 Then
      cbItemText = Cb.list(i)
   Else
      cbItemText = ""
   End If

End Function

Public Sub cbClearSel(ls As ListBox, Optional ByVal bSel As Boolean = 0)
   Dim i As Integer
   
   If ls.Style <> vbListBoxCheckbox Or (bSel = False And ls.SelCount = 0) Or (bSel = True And ls.SelCount = ls.ListCount) Then
      Exit Sub
   End If
   
   For i = 0 To ls.ListCount - 1
      ls.Selected(i) = bSel
   Next i

End Sub

' Copy/ Copia la combo CbSource a la combo CbDest
Sub CboxCpy(CbDest As Control, CbSource As Control, Optional bClean As Boolean = 1)
   Dim i As Integer
   Dim n As Integer

   If bClean Then
      CbDest.Clear ' limpieza
   End If
   
   n = CbSource.ListCount - 1
   If n < 0 Then
      Exit Sub
   End If

   For i = 0 To n
      CbDest.AddItem CbSource.list(i)
      CbDest.ItemData(CbDest.NewIndex) = CbSource.ItemData(i)
   Next i

   CbDest.ListIndex = CbSource.ListIndex
   
End Sub

Public Sub SetCur(ByVal nCur As Long)
   Dim hCur As Long
   Dim hOCur As Long
   
   hCur = LoadCursor(0, nCur)

   hOCur = SetCursor(hCur)

End Sub

Public Sub SetTxRO(Tx As TextBox, ByVal bLocked As Boolean)
   'Dim Rc As Integer

   Tx.Locked = bLocked
   'Rc = SendMessage(Tx.hWnd, EM_SETREADONLY, Bool, 0)

   If bLocked Then
      Tx.MousePointer = vbArrow
      Tx.BackColor = vbButtonFace
      Tx.TabStop = False
   Else
      Tx.MousePointer = vbDefault
      Tx.BackColor = vbWindowBackground
      Tx.TabStop = True
   End If

End Sub
Public Sub SetRO(Tx As TextBox, ByVal bLocked As Boolean)
   Debug.Print "*** Cambiar SetRO por SetTxRO ***"
   Call SetTxRO(Tx, bLocked)
End Sub

Public Function TxLineCount(Tx As TextBox) As Long

   TxLineCount = SendMessage(Tx.hWnd, EM_GETLINECOUNT, 0, 0)

End Function

' Formatea un RUT que viene como "8123456-4" o "8.123.456-4"
Function sFmtRut(ByVal Rut As String) As String
   Dim lRut As Long

   Rut = Trim(Rut)
   
   lRut = Val(ReplaceStr(Rut, ".", ""))

   sFmtRut = ReplaceStr(Format(lRut, NUMFMT), ",", ".") & "-" & UCase(Right(Rut, 1))

End Function

' OJO: El Form debe tener el KeyPreview en True
' Esta función debe ponerse en el evento KeyUp del Form
' Primero revisa si el control activo tiene help, de lo contrario
' asume el del Form.
' Para llamar desde el menu usar: ShowHlp( Me, vbKyF1 )
Public Function ShowHlp(Frm As Form, ByVal KeyCode As Integer, Optional ByVal HelpContextID As Integer = 0) As Integer
   Dim fname As String
   Dim Rc As Long
   Static HtmInfo As ExtInfo_t

   If KeyCode <> vbKeyF1 Or Trim(App.HelpFile) <> "" Then
      Exit Function
   End If

   MsgBox1 "Hlp:[" & App.HelpFile & "]"

   Frm.MousePointer = vbHourglass
   DoEvents

   On Error Resume Next

   If HtmInfo.Ext = "" Then
      ' Vemos con que ve los HTM
      Rc = GetExtInfo(".htm", HtmInfo)
   End If

        If HelpContextID <= 0 Then
                ' Vemos si el control actual tiene help propio
                HelpContextID = Frm.ActiveControl.HelpContextID
                If HelpContextID = 0 Then
                        HelpContextID = Frm.HelpContextID    ' sino, usamos el del Form
                End If
        End If

   If HelpContextID = 0 Then
      fname = "index.htm"
   Else
      fname = GetIniString(gAppPath & "\Help\Help.ini", "Topic", str(HelpContextID), "")
      
      If fname = "" Then
         MsgBox1 "No se encontró información para el tópico solicitado.", vbExclamation
         fname = "index.htm"
      End If
      
   End If

   fname = "file:///" & gAppPath & "/Help/" & fname
   Rc = ShellExecute(Frm.hWnd, "open", HtmInfo.OpenCmd, fname, "", 1)
   If Rc < 32 Then
      MsgBox1 "Error " & Rc & ", " & fname, vbExclamation
   End If

   Frm.MousePointer = vbDefault

End Function

Public Function ShowHelp(Frm As Form, Optional ByVal HelpContextID As Integer = 0, Optional ByVal Url As String = "") As Long
   Dim Rc As Long, Fn As String, Msg As String

   If App.HelpFile = "" Then
      Debug.Print "*** No hay help ***"
      ShowHelp = -1
      Exit Function
   End If

   Fn = App.HelpFile
   If ExistFile(Fn) = False Then
   
      Msg = "No existe el archivo de ayuda" & vbCrLf & Fn
      If Url <> "" Then
         Msg = Msg & vbCrLf & "Puede obtenerlo en " & Url
      End If
   
      MsgBox1 Msg, vbExclamation
      ShowHelp = -2
      Exit Function
   End If

   Rc = ExecCmd("HH.exe """ & Fn & """", vbNormalNoFocus, 0)

   If Rc Then  ' falló
      MsgBox1 "Error " & ERR.LastDllError & " al invocar la ayuda." & vbCrLf & Fn, vbInformation
   End If

End Function

' Content por ejemplo: "Import.htm"
Public Function ShowHelp2(Frm As Form, Optional ByVal Content As String = "") As Long
   Dim Rc As Long, Fn As String, Msg As String, Buf As String, i As Integer, H As Integer

   If App.HelpFile = "" Then
      Debug.Print "*** No hay help ***"
      ShowHelp2 = -1
      Exit Function
   End If
   
   Rc = HtmlHelp(Frm.hWnd, App.HelpFile & IIf(Content <> "", "::/" & Content, ""), HH_DISPLAY_TOPIC, 0)
   If Rc = 0 And StrComp(Left(App.HelpFile, 3), Left(gIniFile, 3), vbTextCompare) Then
      i = rInStr(gIniFile, "\")
      H = rInStr(App.HelpFile, "\")
      
      If H > 0 Then
         Fn = Left(gIniFile, i) & Mid(App.HelpFile, H + 1)
      Else
         Fn = Left(gIniFile, i) & App.HelpFile
      End If
      
      If CopyFile(App.HelpFile, Fn, True) = -1 Then
         Rc = HtmlHelp(Frm.hWnd, Fn & IIf(Content <> "", "::/" & Content, ""), HH_DISPLAY_TOPIC, 0)
      End If
   End If

   ShowHelp2 = Rc
   
End Function

Sub Sleep1(ByVal nSeg As Integer)
   Dim Tm As Double

   Tm = Now + TimeSerial(0, 0, nSeg)

   Do While Tm > Now
      DoEvents
   Loop
       
End Sub
Function SQLNull(ByVal FldName As String) As String

   SQLNull = " (" & FldName & " IS NULL OR " & FldName & "='') "

End Function

Function Trunc(ByVal Buf As String, ByVal wID As Long) As String
   Dim i As Integer

   For i = Len(Buf) To 1 Step -1
      If Printer.TextWidth(Left(Buf, i)) < wID Then
         Trunc = Left(Buf, i)
         Exit Function
      End If
   Next i

   Trunc = "..."

End Function
Sub TxGotFocus(Tx As Control)

   Tx.SelStart = 0
   Tx.SelLength = 64000

End Sub
Sub NumGotFocus(Tx As TextBox)
   Dim Chg As Boolean

   If Trim(Tx) <> "" And Tx.Locked = False Then
      Chg = Tx.DataChanged
      Tx = vFmt(Tx)
      Tx.SelStart = 0
      Tx.SelLength = 32000
      Tx.DataChanged = Chg
   End If

End Sub
Sub NumLostFocus(Tx As TextBox, Optional ByVal nDec As Byte = 0)
   Dim Fmt As String, Chg As Boolean
   
   Chg = Tx.DataChanged
   
   If Trim(Tx) <> "" And Tx.Locked = False Then
   
      If nDec = 0 Then
         Fmt = NUMFMT
      Else
         Fmt = NUMFMT & "." & String(nDec, "0")
      End If
      
      Tx = Format(vFmt(Tx), Fmt)
   ElseIf Trim(Tx) = "" Then
      Tx = ""
   End If
   
   Tx.DataChanged = Chg

End Sub
' Sin Separador de miles
Sub NumLostFocus2(Tx As TextBox, Optional ByVal nDec As Byte = 0)
   Dim Fmt As String, Chg As Boolean
   
   Chg = Tx.DataChanged
   
   If Trim(Tx) <> "" And Tx.Locked = False Then
   
      If nDec = 0 Then
         Fmt = "0"
      Else
         Fmt = "0." & String(nDec, "0")
      End If
      
      Tx = Format(vFmt(Tx), Fmt)
   ElseIf Trim(Tx) = "" Then
      Tx = ""
   End If
   
   Tx.DataChanged = Chg

End Sub
Public Sub RUT_GotFocus(Tx As TextBox)
   Dim Chg As Boolean, lRut As Long

   If Trim(Tx) <> "" And Tx.Locked = False Then
      Chg = Tx.DataChanged
      lRut = vFmtRut(Tx)
      If lRut > 0 Then
         Tx = lRut & "-" & DV_Rut(lRut)
         Tx.SelStart = 0
         Tx.SelLength = 32000
         Tx.DataChanged = Chg
      End If
   End If

End Sub
Public Function RUT_LostFocus(Tx As TextBox, Optional ByVal bSetFocus As Boolean = 0) As Boolean
   Dim lRut As Long, Chg As Boolean
   
   Tx.Text = Trim(Tx.Text)
   RUT_LostFocus = True
   
   If Len(Tx) > 0 And Tx.Locked = False Then
      Chg = Tx.DataChanged
      
      lRut = vFmtRut(Tx)
      If lRut > 0 Then
         Tx = FmtRut(lRut)
      Else
         MsgBox1 "El Rut ingresado es inválido.", vbExclamation
         RUT_LostFocus = False
         
         If bSetFocus Then
            Call Tx.SetFocus
         End If
      End If
      Tx.DataChanged = Chg
   
   End If

End Function
Function ValidId(ByVal VarId As String) As Integer
   Dim i As Integer
   Dim ch As String
   
   ch = UCase(Left(VarId, 1))
   ValidId = False

   If ch < "A" Or ch > "Z" Then
      Exit Function
   End If

   For i = 2 To Len(VarId)
      ch = UCase(Mid(VarId, i, 1))
      
      If ch <> "_" And (ch < "A" Or ch > "Z") And Not IsNumeric(ch) Then
         Exit Function
      End If

   Next i

   ValidId = True

End Function

Public Function MsgValidRut(ByVal Rut As String) As Boolean
   
   MsgValidRut = False
   
   If Rut <> "" Then
      If Not ValidRut(Rut) Then
         MsgBox1 "Rut inválido.", vbOKOnly + vbExclamation
         Exit Function
      End If

      MsgValidRut = True
   End If
   
End Function
Public Function ValidName(TxtName As TextBox, Fld As String, Optional ByVal MsgWarning As Boolean = False) As Boolean
   Dim ch As String
   
   ValidName = False
   TxtName = Trim(TxtName)
   
   If Len(TxtName) < 2 Then
      If MsgWarning Then
      
         If TxtName.DataChanged = False Then ' 25 abr 2017: si es warning y no lo modificó
            ValidName = True
            Exit Function
         End If
            
         If MsgBox1("El largo de " & Fld & " debe ser mayor o igual que 2." & vbCrLf & vbCrLf & "¿Está seguro que desea grabar?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            TxtName.SetFocus
            Exit Function
         ElseIf Len(TxtName) = 0 Then
            ValidName = True
            Exit Function
         End If
      Else
         MsgBox1 "El largo de " & Fld & " debe ser mayor o igual que 2.", vbExclamation
         TxtName.SetFocus
         Exit Function
      End If
   End If
      
   ch = LCase(Left(TxtName, 1))
   
   If (ch < "a" Or ch > "z") And (ch < "à" Or ch > "ü") Then ' 18 dic 2017: se agrega (ch < "à" Or ch > "ü")
      If MsgWarning Then
         If MsgBox1("El primer caracter de " & Fld & " debe ser alfabético." & vbCrLf & vbCrLf & "¿Está seguro que desea grabar?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            TxtName.SetFocus
            Exit Function
         End If
      Else
         MsgBox1 "El primer caracter de " & Fld & " debe ser alfabético.", vbExclamation
         TxtName.SetFocus
         Exit Function
      End If
   End If
      
   ValidName = True
      
End Function

'
' Acepta RUTs del tipo 9999999-V, del tipo 9.999.999-V o 9999999V
' Ver: DV_Rut(), vFmtRut()
Function ValidRut(ByVal Rut As String) As Boolean
   Dim i As Integer, nRut As Long, dv As String

   Rut = UCase(ReplaceStr(Trim(Rut), ".", ""))

   On Error Resume Next
   
   ValidRut = False
   i = InStr(Rut, "-")
      If i Then
      nRut = Val(Left(Rut, i - 1))
      ValidRut = (DV_Rut(nRut) = Mid(Rut, i + 1))
   Else
      nRut = Val(Left(Rut, Len(Rut) - 1))
      ValidRut = (DV_Rut(nRut) = Right(Rut, 1))
   End If
   
   If nRut = 0 Then
      If Trim(Left(Rut, i - 1)) <> "0" Then
         ValidRut = False
      End If
   End If
   
End Function

Function vFmtRut(ByVal Rut As String) As Long
   Dim i As Integer, nRut As Long
   
   If ValidRut(Rut) = False Then
      vFmtRut = 0
      Exit Function
   End If

   Rut = ReplaceStr(Trim(Rut), ".", "")
   i = InStr(Rut, "-")
   
   If i > 0 Then
      vFmtRut = Val(Left(Rut, i - 1))
   Else
      nRut = Val(Left(Rut, Len(Rut) - 1))
      If DV_Rut(nRut) = Right(Rut, 1) Then
         vFmtRut = nRut
      Else
         vFmtRut = Val(Rut)
      End If
   End If

'   Rut = Trim(Rut)
'   Ln = Len(Rut)
'   i = InStr(Rut, "-")
'   On Error Resume Next
'
'   If i = 0 Then
'      i = Ln
'   End If
'
'   If sRut = "" Or Ln <= 2 Or i <= 0 Then
'      vFmtRut = 0
'   Else
'      vFmtRut = vFmt(ReplaceStr(Left(sRut, i - 1), ".", ""))
'   End If

End Function

' Toma un rut y lo deja del tipo nnnnnnn-d o 000nnnnn-d
Function NormRut(ByVal Rut As String, Optional RLen As Integer = 0) As String
   Dim r As Long
   
   Rut = Trim(Rut)
   r = vFmtRut(Rut)
   If r <= 0 Then
      Rut = "0-0"
   Else
      Rut = r & "-" & Right(Rut, 1)
   End If
   
   If RLen > 0 Then
      Rut = Right(String(RLen, "0") & Rut, RLen)
   End If

   NormRut = Rut
   
End Function

' Formatea un RUT que viene como 8123456 ==> 8123456-7
Function FmtRut(ByVal Rut As Double) As String

   FmtRut = ReplaceStr(Format(Rut, NUMFMT), ",", ".") & "-" & DV_Rut(Rut)

End Function

' Calcula el dígito verificador de un RUT
Function DV_Rut(ByVal lRut As Double) As String
   Dim i As Integer
   Dim sRut As String
   Dim inNumeroRut     As Integer
   Dim inMultiplicador As Integer
   Dim inAuxRut        As Integer
   Dim inResto         As Integer

   sRut = Trim(lRut)
   inMultiplicador = 2
   inNumeroRut = 0

   'Se calcula el digito verificador del rut
   For i = Len(sRut) To 1 Step -1
      If inMultiplicador = 8 Then
         inMultiplicador = 2
      End If

      inNumeroRut = inNumeroRut + (inMultiplicador * Val(Mid(sRut, i, 1)))
      inMultiplicador = inMultiplicador + 1
   Next i

   inAuxRut = Int(inNumeroRut / 11)
   inResto = inNumeroRut - (inAuxRut * 11)

   'Retornamos el digito verificador
   If inResto <> 0 Then
      inResto = 11 - inResto
      If inResto = 10 Then
         DV_Rut = "K"
      Else
         DV_Rut = "" & inResto
      End If
   Else
      DV_Rut = "0"
   End If

End Function

' Valida un año y le suma 1900 o 2000 si es necesario
Function ValYear(ByVal Yr As String) As Integer
   Dim yy As Long
   
   yy = Abs(Val(Yr))
   If yy < 0 Or yy > Year(Now) + 50 Then
      yy = 0
      Exit Function
   End If

   If yy < 100 Then
      If yy < AÑO_2DIG Then
         yy = 2000 + yy
      Else
         yy = 1900 + yy
      End If
   ElseIf yy < 1000 Then
      yy = 2000 + yy

   End If
   
   ValYear = yy
   
End Function


Function vFmt(ByVal Buf As String) As Double

   vFmt = 0
   On Error Resume Next
   vFmt = Format(Buf)
   On Error GoTo 0

End Function
' Para ser usado en SQL en los UPDATE, INSERT o SELECT
Function svFmt(ByVal Buf As String) As String

   svFmt = Str0(vFmt(Buf))

End Function

'ffv 25-10-23 veracode 3312124
'comentada por alerta de seguridad
'Function WinExec(ByVal Cmd As String, ByVal ShWnd As Integer) As Long
'   Dim Hnd As Long
'   Dim Tm As Double
'
'   'MsgBox "WinExec=[" & Cmd & "]"
'   Hnd = Shell(Cmd, ShWnd) ' Handle de la ventana DOS
'   If Hnd < 32 Then
'      WinExec = Hnd
'      Exit Function
'   End If
'
'   ' Esperamos a que termine de ejecutar
'   Tm = Now
'   'Do
'      DoEvents
'      If Now - Tm > TimeSerial(0, 10, 0) Then  ' 10 Min
'         MsgBox1 "Por favor revise las ventanas DOS y cierre aquellas que estén finalizadas.", vbExclamation
'         Tm = Now
'      End If
'
'   'Loop While GetModuleUsage(Hnd) > 0
'
'   WinExec = 0
'
'End Function

'ffv 25-10-23 veracode 3312124
'comentada por alerta de seguridad
'Sub WinExec2(ByVal Cmd As String, ByVal ShWnd As Integer)
'   Dim Hnd As Integer
'   Dim Tm As Double
'
'   'msgbox1 "WinExec=[" & Cmd & "]"
'   Hnd = Shell(Cmd, ShWnd) ' Handle de la ventana DOS
'
'   ' Esperamos a que termine de ejecutar
'   Tm = Now
'   'Do
'      DoEvents
'      If Now - Tm > TimeSerial(0, 10, 0) Then  ' 10 Min
'         MsgBox1 "Por favor, revise las ventanas DOS y cierre aquellas que estén finalizadas.", vbExclamation
'         Tm = Now
'      End If
'
'   'Loop While GetModuleUsage(Hnd) > 0
'
'End Sub

' nShowWindow son las constantes SW_...
' nMilliseconds puede ser INFINITE o 0 si no espera nada
Public Function ExecCmd(ByVal Cmdline As String, ByVal nShowWindow As VbAppWinStyle, Optional ByVal nMilliSeconds As Long = INFINITE) As Long
   Dim Proc As PROCESS_INFORMATION
   Dim Start As STARTUPINFO
   Dim Rc As Long, RcW As Long

   ' Initialize the STARTUPINFO structure:
   Start.Cb = Len(Start)
   Start.dwFlags = STARTF_USESHOWWINDOW
   Start.wShowWindow = nShowWindow
   
   ' Start the shelled application:
   Rc = CreateProcessA(0&, Cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc)
   If Rc = 0 Then
      ExecCmd = ERR.LastDllError
      Exit Function
   End If

   Rc = 0
   ' Wait for the shelled application to finish:
   RcW = WaitForSingleObject(Proc.hProcess, nMilliSeconds)
   Call GetExitCodeProcess(Proc.hProcess, RcW)
   Call CloseHandle(Proc.hProcess)
   
   ExecCmd = Rc
   
End Function

Function D2D(Dat As Double)
   D2D = Dat
End Function
Function Log2(ByVal Num As Double) As Double
   Log2 = Log(Num) / Log(2)
End Function
' Redondea un número con la cantidad de decimales indicados
Public Function Round(ByVal Valor As Double, Optional ByVal nDec As Byte = 0) As Double
   Dim i As Integer
   Dim v As Double
   
   If Valor = 0 Then
      Round = 0
   
   ElseIf nDec = 0 Then
      Round = Int(Abs(Valor) + 0.5000000000001) * Sgn(Valor)
   
   ElseIf nDec < 15 Then
   
      v = Abs(Valor) * (10 ^ nDec)
      Round = Int(v + 0.5000000000001) / (10 ^ nDec) * Sgn(Valor)
      
   Else
      Round = Valor
   
   End If
   
End Function

Public Function MsgBox1(ByVal Msg As String, Optional ByVal Flags As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
   Dim Sound As Long
      
   Sound = Flags And (vbExclamation Or vbCritical Or vbInformation)

   If gMsgBoxLog Then
      Call AddLog(Msg)
      MsgBox1 = vbNo
   Else
      MsgBeep Sound
      MsgBox1 = MsgBox(Msg, Flags)
   End If

End Function

Public Sub MsgErr(ByVal Msg As String, Optional ByVal Flags As VbMsgBoxStyle = vbExclamation)
   Dim Buf As String
   
   If Msg <> "" Then
      Msg = NL & Msg
   End If
      
   MsgBeep vbExclamation
   Buf = "Error " & ERR.Number & ", " & vbErrMsg(ERR.Number, ERR.Description)
   If ERR.LastDllError Then
      Buf = Buf & vbCrLf & " (DllErr=" & ERR.LastDllError & ")"
   End If
   
   MsgBox Buf & Msg, Flags
   
End Sub


Public Sub PostClick(Button As CommandButton, Optional bChkEnabled As Boolean = 1)
   Dim Rc As Long

   If bChkEnabled = False Or (Button.Enabled = True And Button.visible = True) Then
      Rc = PostMessage(Button.hWnd, WM_LBUTTONDOWN, 0, 0)
      Rc = PostMessage(Button.hWnd, WM_LBUTTONUP, 0, 0)
   End If
   
End Sub
' Cuenta los días LMMJV entre dos fechas
Public Function DiasHabiles(ByVal Fecha1 As Long, ByVal Fecha2 As Long) As Integer
   Dim l As Long
   Dim d As Integer
   
   For l = Fecha1 To Fecha2
      If Weekday(l, vbMonday) < 6 Then
         d = d + 1
      End If
   Next l

   DiasHabiles = d

End Function
Public Function NextDiaHabil(ByVal Fecha As Long) As Long
   Dim d As Integer, W As Integer
   
   For d = 1 To 10
      W = Weekday(Fecha + d, vbMonday)
      If W >= 1 And W <= 5 Then
         NextDiaHabil = Fecha + d
         Exit For
      End If
   Next d

End Function



' Funcion para leer archivos Excel sep. por tabs
' con función Line Input
Public Function NextField(Buf As String, Optional ByVal Sep As String = vbTab) As String
   Dim i As Long

   i = InStr(Buf, Sep)
   If i Then
      NextField = Trim(Left(Buf, i - 1))
      
      Buf = Mid(Buf, i + Len(Sep))
   Else
      NextField = Trim(Buf)
      Buf = ""
   
   End If

End Function

Public Function NextField2(ByVal Buf As String, Pos As Long, Optional ByVal Sep As String = vbTab) As String
   Dim i As Long, l As Long
   
   l = Len(Buf)
   If Pos > l Then
      NextField2 = vbNullString
      Exit Function
   End If
   
   i = InStr(Pos, Buf, Sep, vbBinaryCompare)
   
   If i Then
      NextField2 = Mid(Buf, Pos, i - Pos)
      Pos = i + Len(Sep)
   Else
      NextField2 = Mid(Buf, Pos)
      Pos = l + Len(Sep)
   End If
      
End Function
Public Function NextCsv_old(ByVal Buf As String, Pos As Long, Optional ByVal Sep As String = vbTab) As String
   Dim i As Long, l As Long
   
   l = Len(Buf)
   If Pos > l Then
      NextCsv_old = ""
      Exit Function
   End If
   
   If Mid(Buf, Pos, 1) = """" Then ' hay que buscar la comilla que cierra el campo
   
      i = InStr(Pos, Buf, """" & Sep, vbBinaryCompare)
      
      If i > 0 Then
         NextCsv_old = Mid(Buf, Pos + 1, i - Pos - 1)
         Pos = i + 2
      Else
         NextCsv_old = Mid(Buf, Pos + 1, l - Pos)
         Pos = l + 2
      End If

   Else
      i = InStr(Pos, Buf, Sep, vbBinaryCompare)
      
      If i Then
         NextCsv_old = Mid(Buf, Pos, i - Pos)
         Pos = i + 1
      Else
         NextCsv_old = Mid(Buf, Pos)
         Pos = l + 1
      End If
   End If
      
End Function

Public Function NextCsv(ByVal Buf As String, Pos As Long, Optional ByVal Sep As String = vbTab) As String
   Dim i As Long, l As Long
   
   l = Len(Buf)
   If Pos > l Then
      NextCsv = ""
      Exit Function
   End If
   
   If Mid(Buf, Pos, 1) = """" Then
   
      i = InStr(Pos, Buf, """" & Sep, vbBinaryCompare)
   
      If i > 0 Then
         NextCsv = Mid(Buf, Pos + 1, i - Pos - 1)
         Pos = i + 2
      ElseIf Mid(Buf, l, 1) = """" Then
         NextCsv = Mid(Buf, Pos + 1)
         Pos = l + 2
      Else
         NextCsv = vbNullString
      End If
   
   Else
      i = InStr(Pos, Buf, Sep, vbBinaryCompare)
      
      If i > 0 Then
         NextCsv = Mid(Buf, Pos, i - Pos)
         Pos = i + 1
      Else
         NextCsv = Mid(Buf, Pos)
         Pos = l + 1
      End If
   End If
      
End Function

' Verifica si el archivo existe o no
Public Function ExistFile(ByVal fname As String) As Boolean

   On Error Resume Next
   
   ExistFile = False
   
   If fname = "" Then
      Exit Function
   End If
   
   'ExistFile = (Dir(FName) <> "")
   ExistFile = (FileLen(fname) >= 0)

End Function
' Elimina un archivo
Public Function RemoveFile(ByVal fname As String) As Long

   On Error Resume Next
   
   Kill fname
   
   RemoveFile = ERR.Number

End Function

Public Sub FrameEnable(fr As Frame, ByVal bool As Boolean)
   Dim Cnt As Control
   Dim Frm As Form

   fr.Enabled = bool

   On Error Resume Next
   
   Set Frm = fr.Parent

   For Each Cnt In Frm.Controls
      If Cnt.Container.hWnd = fr.hWnd Then
         'If TypeOf Cnt Is TextBox Then
         '   Call SetTxRO(Cnt, Not Bool)
         'Else
            Cnt.Enabled = bool
         'End If
      End If
   
   Next Cnt

End Sub
#If (Pamopt And 1) Then
Public Sub BtFechaImg(bt As CommandButton)

   If gFrmMain Is Nothing Then
      Debug.Print "*** Falta asignar el gFrmMain ***"
      Exit Sub
   End If

   If bt.Style = vbButtonGraphical Then
      On Error Resume Next
   
      bt.UseMaskColor = True
      bt.MaskColor = &HFFFFFF
      bt.Picture = gFrmMain.Im_Down
      If bt.Picture <> 0 Then
         bt.Caption = ""
      End If
   End If
   
End Sub
#End If

' hh:nn:ss o hh:nn o h
Public Function vFmtHour(ByVal Hora As String) As Double
   Dim H As Integer, m As Integer, s As Integer, i As Integer, x As Integer, k As Integer
   
   If Trim(Hora) = "" Then
      vFmtHour = -1
      Exit Function
   End If
   
   For k = 1 To 3
   
      i = InStr(Hora, ":")
      If i > 3 Then
         vFmtHour = 0
         Exit Function
      ElseIf i <= 0 Then ' h
         x = Val(Hora)
      Else
         x = Val(Left(Hora, i - 1))
         Hora = Mid(Hora, i + 1)
      End If
   
      Select Case k
         Case 1:
            H = x
         Case 2:
            m = x
         Case 3:
            s = x
      End Select
         
      If i <= 0 Then
         Exit For
      End If
   Next k
   
   vFmtHour = TimeSerial(H, m, s)

End Function

' hhnn o hnn o nn
Public Function vFmtHour2(ByVal Hora As Integer) As Double
   Dim H As Integer, m As Integer, i As Integer, x As Integer, k As Integer
   
   m = Hora Mod 100
   H = Hora \ 100
      
   vFmtHour2 = TimeSerial(H, m, 0)

End Function
' hhnn o hnn o nn
Public Function vFmtHour3(ByVal Hora As Integer) As Double
   Dim H As Integer, m As Integer, i As Integer, x As Integer, k As Integer
   
   m = Hora Mod 100
   H = Hora \ 100
      
   vFmtHour3 = H + m / 60#

End Function

Public Function GetLastUser(ByVal IniFile As String, Optional ByVal DefUser As String = "", Optional Section As String = "") As String
   Dim Rc As Long
   Dim Buff As String * 50

   If IniFile = "" Then
      MsgBox1 "Falta definir el .INI", vbExclamation
   End If
   
   If Section = "" Then
      Section = "Config-" & W.UserName
   End If
   
   GetLastUser = GetIniString(IniFile, Section, "LastUser", DefUser)
   
'   Rc = GetPrivateProfileString(Section, "LastUser", DefUser, Buff, 40, IniFile)
'   GetLastUser = Left(Buff, Rc)

End Function
Public Sub SetLastUser(ByVal IniFile As String, ByVal UserName As String, Optional Section As String = "")
   Dim Rc As Long

   If Section = "" Then
      Section = "Config-" & W.UserName
   End If

   Rc = WritePrivateProfileString(Section, "LastUser", UserName, IniFile)

End Sub

' Formatea un Citizen ID que viene como 8123456
Function FmtCID(ByVal CID As String, Optional ByVal bForceRUT As Boolean = 1) As String
   Dim lRut As Long
   
   CID = Trim(CID)
   
   If gValidRut And bForceRUT Then
      lRut = Val(ReplaceStr(CID, ".", ""))
      FmtCID = ReplaceStr(Format(lRut, NUMFMT), ",", ".") & "-" & DV_Rut(lRut)
   Else
      FmtCID = CID
   End If

End Function

' Formatea un CID que viene como "8123456-4" o "8.123.456-4"
Function sFmtCID(ByVal CID As String, Optional ByVal bForceRUT As Boolean = 1) As String
   Dim lRut As Long

   CID = Trim(CID)

   If gValidRut And bForceRUT Then

      lRut = Val(ReplaceStr(CID, ".", ""))

      sFmtCID = ReplaceStr(Format(lRut, NUMFMT), ",", ".") & "-" & UCase(Right(lRut, 1))
   Else
      sFmtCID = CID
   End If

End Function

Public Function MsgValidCID(ByVal CID As String, Optional ByVal bForceRUT As Boolean = 1) As Boolean
   
   If Not ValidCID(CID, bForceRUT) Then
      If gValidRut And bForceRUT Then
         MsgBox1 "RUT inválido.", vbExclamation
      Else
         MsgBox1 "Identificación inválida.", vbExclamation
      End If
      MsgValidCID = False
      Exit Function
   End If

   MsgValidCID = True
   
End Function

'
' Acepta RUTs del tipo 9999999-V y del tipo 9.999.999-V
'
Function ValidCID(ByVal CID As String, Optional ByVal bForceRUT As Boolean = 1) As Boolean

   CID = Trim(CID)

   If gValidRut And bForceRUT Then
      
      On Error Resume Next
      ValidCID = ValidRut(CID)
   Else
      ValidCID = (CID <> "" And CID <> "0")
   End If

End Function

Function vFmtCID(ByVal CID As String, Optional ByVal bForceRUT As Boolean = 1) As String
   Dim Ln As Integer

   CID = Trim(CID)
   If CID = "" Then
      vFmtCID = ""
      Exit Function
   End If

   If gValidRut And bForceRUT Then
      vFmtCID = vFmtRut(CID)
   Else
      vFmtCID = CID
   End If

End Function

Public Function LTrim0(ByVal Cod As String) As String
    Dim i As Integer

   For i = 1 To Len(Cod)
      If Mid(Cod, i, 1) <> "0" Then
         LTrim0 = Mid(Cod, i)
         Exit Function
      End If
   Next i

   LTrim0 = ""

End Function
Public Function L0Trim(ByVal Cod As String) As String
   L0Trim = LTrim0(Cod)
End Function

Public Function RTrimLF(ByVal Buf As String) As String
   Dim i As Integer, l As Integer

   Buf = Trim(Buf)
   l = Len(Buf)

   Do While l > 0

      If l > 1 Then
         If Mid(Buf, l - 1, 2) = vbCrLf Then
            l = l - 2
         End If
      End If
      
      If l > 0 Then
         If Mid(Buf, l, 1) = vbLf Then
            l = l - 1
         ElseIf Mid(Buf, l, 1) = " " Then
            l = l - 1
         Else
            Exit Do
         End If
      End If

   Loop

   RTrimLF = Left(Buf, l)

End Function

' Para los decimales queden como 0.ddddd y no como .ddddd
Public Function Str0(ByVal Num As Double) As String

   Str0 = ReplaceStr(CStr(Num), ",", ".")

   'If Abs(Num) >= 1 Then
   '   Str0 = Str(Num)
   'ElseIf Num > 0 Then
   '   Str0 = "0" & Trim(Str(Num))
   'ElseIf Num < 0 Then
   '   Str0 = "-0" & Trim(Str(Abs(Num)))
   'Else
   '   Str0 = "0"
   'End If

End Function

' *** ANTIGUA *** ahora usar GetExtInfo
Public Function GetOpenCmd(ByVal Ext As String) As String
   Dim ExtInfo As ExtInfo_t
   
   If GetExtInfo(Ext, ExtInfo) Then
      GetOpenCmd = ExtInfo.OpenCmd
   Else
      GetOpenCmd = ""
   End If
   
End Function

' Busca el comando para abrir un documento con extensión Ext (ej. ".rtf", ".xls", ...)
' Para generar el comando usar la GenCmd por que puede venir un %1
Public Function GetExtInfo(ByVal Ext As String, ExtInfo As ExtInfo_t) As Boolean
   Dim Path As String, Path0 As String, Buf As String, Buf1 As String
   Dim Rc As Long, Key As Long, BLen As Long

   GetExtInfo = False
   ExtInfo.Ext = ""
   ExtInfo.OpenCmd = ""
   ExtInfo.DdeOpen = ""
   ExtInfo.DdeOpenApp = ""
   ExtInfo.DdeOpenTopic = ""
   ExtInfo.PrintCmd = ""
   ExtInfo.DdePrint = ""
   ExtInfo.DdePrtApp = ""
   ExtInfo.DdePrtTopic = ""
   ExtInfo.bReg = False

   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Ext, 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Path0 = Trim(FwLeft(Buf1, BLen - 1))   ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Call AddDebug("GetExtInfo: Path0=[" & Path0 & "]")

   Rc = RegCloseKey(Key)
   Key = 0
   
   If Path0 = "" Then
      Exit Function
   End If
   
   ExtInfo.Path = Path0
   
   '********** OPEN **********
   
   Path = Path0 & "\Shell\Open"

   Call AddDebug("GetExtInfo: Path=[" & Path & "]")

   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\Command", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Call AddDebug("GetExtInfo: Cmd=[" & Buf & "]")

   Rc = RegCloseKey(Key)

   If Left(Buf, 1) = """" Then
      Rc = InStr(2, Buf, """")
      Buf = Mid(Buf, 2, Rc - 2)
   'Else
   '   Rc = InStr(Buf, " ")
   '   Buf = Left(Buf, Rc - 1)
   End If

   ExtInfo.bReg = True
   ExtInfo.Ext = Ext
   ExtInfo.OpenCmd = ConvSysVars(Buf)
   GetExtInfo = True
   
   Call AddDebug("GetExtInfo: OpenCmd=[" & ExtInfo.OpenCmd & "]")
   
   ' Vemos si tiene DDE
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3145")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Call AddDebug("GetExtInfo: 3149")
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Call AddDebug("GetExtInfo: 3151 len=" & Len(Buf1) & ", BLen=" & BLen)
   
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Call AddDebug("GetExtInfo: 3154")

   Rc = RegCloseKey(Key)

   ExtInfo.DdeOpen = Buf

   ' Vemos si tiene DDE-Application
   Call AddDebug("GetExtInfo: 3161")
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec\Application", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3164")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   ExtInfo.DdeOpenApp = Buf

   ' Vemos si tiene DDE-Topic
   Call AddDebug("GetExtInfo: 3176")
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec\Topic", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3183")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   ExtInfo.DdeOpenTopic = Buf

   '********** PRINT **********
   
   Call AddDebug("GetExtInfo: 3196")
   
   Path = Path0 & "\Shell\Print"

   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\Command", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3205")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   'If Left(Buf, 1) = """" Then
      'Rc = InStr(2, Buf, """")
    '  Buf = Mid(Buf, 2, Len(Buf) - 1)
   'Else
   '   Rc = InStr(Buf, " ")
   '   Buf = Left(Buf, Rc - 1)
   'End If

   Call AddDebug("GetExtInfo: 3223")

   ExtInfo.Ext = Ext
   ExtInfo.PrintCmd = ConvSysVars(Buf)
   
   GetExtInfo = True
   
   Call AddDebug("GetExtInfo: 3229")

   ' Vemos si tiene DDE
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3238")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   ExtInfo.DdePrint = Buf

   ' Vemos si tiene DDE-Application
   Call AddDebug("GetExtInfo: 3250")
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec\Application", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   ExtInfo.DdePrtApp = Buf

   ' Vemos si tiene DDE-Topic
   Call AddDebug("GetExtInfo: 3267")
   
   Rc = RegOpenKeyEx(HKEY_CLASSES_ROOT, Path & "\DdeExec\Topic", 0, KEY_QUERY_VALUE, Key)
   If Rc <> ERROR_SUCCESS Then
      Exit Function
   End If

   Call AddDebug("GetExtInfo: 3274")

   BLen = 200
   Buf1 = Space(BLen + 10)
   Rc = RegQueryValueExS(Key, "", 0, REG_SZ, Buf1, BLen)
   Buf = Trim(FwLeft(Buf1, BLen - 1)) ' 24 oct 2014: pam: usa FwLeft porque a veces retorna BLen=0

   Rc = RegCloseKey(Key)

   Call AddDebug("GetExtInfo: 3283")

   ExtInfo.DdePrtTopic = Buf

End Function
Public Function GenCmd(ExtInfo As ExtInfo_t, ByVal Oper As String, ByVal fname As String) As String
   Dim Cmd As String, i As Integer

   Select Case UCase(Left(Oper, 1))
      Case "O":   ' Open
         Cmd = ExtInfo.OpenCmd
      Case "P":   ' Print
         Cmd = ExtInfo.PrintCmd
      Case Else:
         Cmd = ""
         Exit Function
   End Select
      
   If fname <> "" Then
      If InStr(Cmd, "%") Then
         Cmd = ReplaceStr(Cmd, "%1", fname)
      Else
         i = InStr(1, Cmd, " /dde", vbTextCompare)
         If i Then
            If i + 4 = Len(Cmd) Then ' está al final
               Cmd = Left(Cmd, i)
            Else
               Cmd = ReplaceStr(Cmd, " /dde ", "")
            End If
         End If
         
         ' lo ponemos entre " por si acaso
         Cmd = """" & Cmd & """ """ & fname & """"
      End If
   
   End If

   GenCmd = Cmd

End Function

Public Function MinToolTip(Frm As Form, ByVal Text As String, ByVal wID As Integer)

   If Frm.TextWidth(Text) * 1.05 > wID Then
      MinToolTip = Text
   Else
      MinToolTip = ""
   End If

End Function
'*** 9 MAY 2005 PAM - Recibe control DriveListBox para determinar path absoluto de unidades mapeadas
' vemos si es una unidad de Red y ubicamos su mapeo real
Public Function GetAbsPath(ByVal Path As String, Drv As DriveListBox) As String
   Dim i As Integer, j As Integer, k As Integer, Aux As String, DrvPath As String

   Path = Trim(Path)

   If Mid(Path, 2, 1) = ":" Then
      DrvPath = Left(Path, 2) 'H:
      For i = 0 To Drv.ListCount - 1
         If StrComp(DrvPath, Left(Drv.list(i), 2)) = 0 Then
            Aux = Drv.list(i)  ' K: [\\server\dir1]
            j = InStr(Aux, "[")
            k = InStr(Aux, "]")
            If j <> 0 And k <> 0 Then
               Aux = Mid(Aux, j + 1, k - j - 1)
               If Left(Aux, 2) = "\\" Then
                  Path = ReplaceStr(Path, Left(Path, 2), Aux)
               End If
            End If
            
            Exit For
         End If
      Next i
   End If

   GetAbsPath = Path

End Function

' Convierte un path en relativo al AppPath. Funciona con MkAbsPath
Public Function MkRelPath(ByVal Path As String) As String
   Dim LAP As Integer
   
   LAP = Len(W.AppPath)
   
   If StrComp(Left(Path, LAP), W.AppPath, vbTextCompare) = 0 Then
      MkRelPath = "$(AppPath)" & Mid(Path, LAP + 1)
   Else
      MkRelPath = Path
   End If

End Function
' Convierte un path en absoluto. Funciona con MkRelPath
Public Function MkAbsPath(ByVal Path As String) As String
   
   Path = ReplaceStr(Path, "$(AppPath)", W.AppPath)
   MkAbsPath = ReplaceStr(Path, "(AppPath)", W.AppPath)
   
End Function

Public Function SendEmail(ByVal hWnd As Long, ByVal ToEmail As String, ByVal ToName As String, ByVal Subject As String, Optional ByVal Body As String = "", Optional ByVal cc As String = "") As Boolean
   Dim i As Integer
   Dim Buf As String
   Dim Rc As Long

   SendEmail = False

   ToEmail = Trim(ToEmail)
   If ToEmail = "" Then
      Exit Function
   End If

   i = InStr(ToEmail, "@")
   If i = 0 Or InStr(i + 1, ToEmail, ".", vbBinaryCompare) = 0 Then
      Exit Function
   End If
      
   Subject = Trim(Subject)
    
   Buf = ""
   If Subject <> "" Then
      Subject = ReplaceStr(Subject, " ", "%" & Hex(Asc(" ")))
      Subject = ReplaceStr(Subject, "?", "%" & Hex(Asc("?")))
      Subject = ReplaceStr(Subject, "&", "%" & Hex(Asc("&")))

      Buf = "?Subject=" & Subject
   End If
   
   cc = Trim(cc)
   If cc <> "" Then
      cc = ReplaceStr(cc, " ", "%" & Hex(Asc(" ")))
      cc = ReplaceStr(cc, "?", "%" & Hex(Asc("?")))
      cc = ReplaceStr(cc, "&", "%" & Hex(Asc("&")))

      If Buf = "" Then
         Buf = "?"
      Else
         Buf = Buf & "&"
      End If

      Buf = Buf & "CC=" & cc
   
   End If
   
   Body = Trim(Body)
   
   If Body <> "" Then
      
      Body = ReplaceStr(Body, " ", "%" & Hex(Asc(" ")))
      Body = ReplaceStr(Body, "?", "%" & Hex(Asc("?")))
      Body = ReplaceStr(Body, "&", "%" & Hex(Asc("&")))
      Body = ReplaceStr(Body, vbCr, "%" & Hex2(Asc(vbCr), 2))
      Body = ReplaceStr(Body, vbLf, "%" & Hex2(Asc(vbLf), 2))
      
      If Buf = "" Then
         Buf = "?"
      Else
         Buf = Buf & "&"
      End If
   
      Buf = Buf & "Body=" & Body
   End If
   
   ToName = Trim(ToName)
   If ToName <> "" Then
      ToName = ReplaceStr(ToName, " ", "%" & Hex(Asc(" ")))
   End If

   ' si tiene , o ; no se puede poner el nombre
   If InStr(ToEmail, ",") = 0 And InStr(ToEmail, ";") = 0 Then
      ToEmail = ToName & "<" & ToEmail & ">"
   End If

   Buf = "mailto:" & ToEmail & Buf
   
   Rc = ShellExecute(hWnd, "open", Buf, "", "", 1)
   
   SendEmail = (Rc >= 32)

End Function

Public Sub CloseObj(obj As Object)

   If Not obj Is Nothing Then
      obj.Close
      Set obj = Nothing
   End If

End Sub

' Retorna True si el bit indicado está en uno.
' BitPos debe ir entre 0 y 31
Function ChkBit(ByVal Word As Long, ByVal BitPos As Byte) As Boolean

        ChkBit = (Word And (2 ^ BitPos))

End Function


'Sirve para retornar el valor de una option, si tiene más de un indice
' y estan definidos por Const y se inician de cero
Public Function ValueOption(Frm As Form, ByVal OpName As String) As Integer
   Dim i As Integer
   Dim TName As String
   Dim Opt As Control
   
   ValueOption = -1
   
   For i = 0 To Frm.Controls.Count - 1
   
      TName = TypeName(Frm.Controls(i))
   
      'If TypeOf Opt Is OptionButton Then
      If TName = "OptionButton" Then
         Set Opt = Frm.Controls(i)
         If StrComp(Opt.Name, OpName, vbTextCompare) = 0 And Opt.Value = True Then
            ValueOption = Opt.Index
            Exit For
            
         End If
      End If
   Next i

End Function
' Calcula cuantos meses (con decimales) hay entre dos fechas
Public Function MonthDiff(ByVal Fecha1 As Double, ByVal Fecha2 As Double) As Single
        Dim F1 As Long, F2 As Long
        Dim D1 As Integer, D2 As Integer

        F1 = Int(Fecha1)
        F2 = Int(Fecha2)

        D1 = Day(F1)
        D2 = Day(F2)

        MonthDiff = DateDiff("m", F1, F2) + (D2 - D1) / 30

End Function

Function Ceil(ByVal Number As Double) As Double

   Ceil = Int(Number + 0.999999999999)

End Function

Public Function StrLen(ByVal Buf As String) As Long
   Dim l As Long
   
   l = InStr(Buf, Chr(0))
   If l Then
      StrLen = l - 1
   Else
      StrLen = Len(Buf)
   End If

End Function
Public Function Trim0(ByVal Buf As String) As String
   Dim l As Long
   
   l = InStr(Buf, Chr(0))
   If l Then
      Trim0 = Left(Buf, l - 1)
   Else
      Trim0 = Buf
   End If

End Function

Public Function ParaUrl(ByVal Buf As String) As String
   Dim i As Long
   Dim Out As String, ch As String * 1
   
   Out = ""
   For i = 1 To Len(Buf)
      ch = Mid(Buf, i, 1)
      If IsNumeric(ch) Or UCase(ch) <> LCase(ch) Then
         Out = Out & ch
      Else
         Out = Out & "%" & Right("0" & Hex(Asc(ch)), 2)
      End If
   Next i

   ParaUrl = Out

End Function

Public Function DeUrl(ByVal Buf As String) As String
   Dim i As Long, j As Long
   Dim Out As String, Hx As String, ch As String * 1
   
   j = 1
   Out = ""
   Do
      i = InStr(j, Buf, "%", vbBinaryCompare)
      If i Then
      
         Hx = Mid(Buf, i + 1, 2)
           
         ch = Chr("&H" & Hx)
         
         Out = Out & Mid(Buf, j, i - j) & ch
         
         j = i + 3
           
      Else
         Out = Out & Mid(Buf, j)
         Exit Do
      End If
   
   Loop
   
   DeUrl = Out

End Function


Public Function TestWriteFile(ByVal fname As String) As Integer
   Dim Fd As Long
   Dim Exist As Boolean
   
   On Error Resume Next
   
   If Dir(fname) <> "" Then
      Exist = True
   Else
      Exist = False
   End If
   
   Fd = FreeFile()
   
   Open fname For Append As #Fd
   
   TestWriteFile = ERR
   
   If ERR = 0 Then
      Close #Fd
      
      If Exist = False Then ' Si no existía lo borramos porque el Append lo crea
         Kill fname
      End If
   End If

End Function
' Ve si hay permiso de escritura en el directorio
Public Function TestWriteDir(ByVal Fn As String, ByVal bMsg As Boolean) As Boolean
   Dim i As Integer, Fd As Long, Path As String
   
   If Right(Fn, 1) <> "\" Then
      i = InStrRev(Fn, "\")
      If i > 0 Then
         Path = Left(Fn, i)
      End If
   Else
      Path = Fn
   End If

   Fn = Path & "test.txt"
   
   ERR.Clear
   
   On Error Resume Next
   Fd = FreeFile()
   
   Open Fn For Output As #Fd
   Print #Fd, "Test"
   Close #Fd
   
   TestWriteDir = (ERR.Number = 0)

   If ERR.Number And bMsg Then
      MsgBox1 "Error " & ERR.Number & " al escribir en el directorio" & vbCrLf & Path & vbCrLf & ERR.Description, vbExclamation
   End If
   
   DelFile Fn

End Function
Function Frac(ByVal Number As Double) As Double

   Frac = Number - Int(Number)

End Function


Public Sub AddTxt(Txt As TextBox, ByVal Msg As String, Optional ToLog As Boolean = False)

   If Txt = "" Then
      Txt = Msg
   Else
      Txt = Txt & vbCrLf & Msg
   End If

   If ToLog Then
      Call AddLog(Msg)
   End If

   Txt.SelStart = Len(Txt.Text)
   DoEvents

End Sub

' Crea ShortCuts en el Desktop, StartMenu y Programs
' LnkFile: "$Desktop\nombre"
' LnkFile: "$StartMenu\nombre"
' LnkFile: "$Programs\nombre"
Public Sub CreateLnk(ByVal LnkFName As String, Optional ByVal CmdFName As String = "", Optional ByVal IconFName As String = "", Optional ByVal IconIndex As Integer = 0)
   Dim LnkPath As String, Buf As String, Buff As String * 201, UserDir As String
   Dim Fd As Long, Rc As Long, hKey As Long

   On Error Resume Next

   UserDir = Environ("USERPROFILE")
   If UserDir = "" Then
      UserDir = Environ("windir")
   End If
   
   UserDir = Trim(UserDir)
 
   If UserDir = "" Then
      Exit Sub
   End If
   
   If InStr(1, LnkFName, "$Desktop", vbTextCompare) Then
      If Dir(UserDir & "\Escritorio\*.*") <> "" Then
         LnkPath = UserDir & "\Escritorio"
      ElseIf Dir(UserDir & "\Desktop\*.*") <> "" Then
         LnkPath = UserDir & "\Desktop"
      Else
         Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, hKey)
         If Rc = 0 Then
            Rc = RegQueryValueExS(hKey, "Desktop", 0, REG_SZ, Buff, 200)
            LnkPath = Left(Buff, StrLen(Buff))
            Rc = RegCloseKey(hKey)
         Else
            Exit Sub
         End If
      End If
      
      LnkFName = ReplaceStr(LnkFName, "$Desktop", LnkPath) & ".url"
      
   ElseIf InStr(1, LnkFName, "$StartMenu", vbTextCompare) Then
      If Dir(UserDir & "\Menu Inicio\*.*") <> "" Then
         LnkPath = UserDir & "\Menu Inicio"
      ElseIf Dir(UserDir & "\Start Menu\*.*") <> "" Then
         LnkPath = UserDir & "\Start Menu"
      Else
         Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, hKey)
         If Rc = 0 Then
            Rc = RegQueryValueExS(hKey, "Start Menu", 0, REG_SZ, Buff, 200)
            LnkPath = Left(Buff, StrLen(Buff))
            Rc = RegCloseKey(hKey)
         Else
            Exit Sub
         End If
      End If
      
      LnkFName = ReplaceStr(LnkFName, "$StartMenu", LnkPath) & ".url"

   ElseIf InStr(1, LnkFName, "$Programs", vbTextCompare) Then
      If Dir(UserDir & "\Menu Inicio\Programas\*.*") <> "" Then
         LnkPath = UserDir & "\Menu Inicio\Programas"
      ElseIf Dir(UserDir & "\Start Menu\Programs\*.*") <> "" Then
         LnkPath = UserDir & "\Start Menu\Programs"
      Else
         Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, hKey)
         If Rc = 0 Then
            Rc = RegQueryValueExS(hKey, "Programs", 0, REG_SZ, Buff, 200)
            LnkPath = Left(Buff, StrLen(Buff))
            Rc = RegCloseKey(hKey)
         Else
            Exit Sub
         End If
      End If
      
      LnkFName = ReplaceStr(LnkFName, "$Programs", LnkPath) & ".url"

   ElseIf InStr(1, LnkFName, "$Startup", vbTextCompare) Then
      If Dir(UserDir & "\Menu Inicio\Programas\Inicio\*.*") <> "" Then
         LnkPath = UserDir & "\Menu Inicio\Programas\Inicio"
      ElseIf Dir(UserDir & "\Start Menu\Programs\*.*") <> "" Then
         LnkPath = UserDir & "\Start Menu\Programs\Startup"
      Else
         Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", 0, KEY_READ, hKey)
         If Rc = 0 Then
            Rc = RegQueryValueExS(hKey, "Startup", 0, REG_SZ, Buff, 200)
            LnkPath = Left(Buff, StrLen(Buff))
            Rc = RegCloseKey(hKey)
         Else
            Exit Sub
         End If
      End If
      
      LnkFName = ReplaceStr(LnkFName, "$Startup", LnkPath) & ".url"
   End If

   If CmdFName = "" Then
      'Rc = GetModuleFileName(0, Buff, 200)
      CmdFName = gAppPath & "\" & App.EXEName & ".exe"
      'CmdFName = Left(Buff, Rc)
   End If
   
   If Trim(IconFName) = "" Then
      IconFName = CmdFName
   End If
      
   Fd = FreeFile()
   Open LnkFName For Binary As #Fd
   If ERR = 0 Then
      
      Buf = "[InternetShortcut]" & vbCrLf
      Buf = Buf & "URL=" & CmdFName & vbCrLf
      Buf = Buf & "IconIndex=" & IconIndex & vbCrLf
      Buf = Buf & "IconFile=" & IconFName & vbCrLf
      Put #Fd, 1, Buf
      Close #Fd
   End If
   
End Sub

Public Sub CreateShortcut(Optional ByVal FilePath As String = "", Optional ByVal ShortcutName As String = "")
   Dim Filesys As Object ' FileSystemObject - References: Microsoft Scripting Runtime
   Dim WshShell As Object
   Dim oShellLink As Object
   Dim DesktopPath As String
   
   If FilePath = "" Then
      FilePath = App.Path & "\" & App.EXEName & ".exe"
   End If
   
   If ShortcutName = "" Then
      ShortcutName = App.Title
   End If
   
   Set Filesys = CreateObject("Scripting.FileSystemObject")
   
   Set WshShell = CreateObject("WScript.Shell")
   DesktopPath = WshShell.SpecialFolders("Desktop")
   
   Set oShellLink = WshShell.CreateShortcut(DesktopPath & "\" & ShortcutName & ".lnk")
   If Filesys.FileExists(oShellLink) Then
      MsgBox1 "Ya existe un ícono con el nombre '" & ShortcutName & "' en el escritorio." & vbCrLf & "Elimínelo y vuelva a intentar.", vbExclamation
   Else
       
      oShellLink.TargetPath = FilePath
      oShellLink.IconLocation = FilePath
      oShellLink.WorkingDirectory = FilePath
      oShellLink.Save
   End If
   
   Set oShellLink = Nothing
   Set WshShell = Nothing
   Set Filesys = Nothing
   
End Sub

' Se asume que vienen dos caracteres: A0, 1B, ...
Public Function Hex2Char(ByVal HexValue As String) As String
   Const HEXCHARS = "0123456789ABCDEF"
   Dim p As Integer, i As Integer, l As Integer, Val As Integer
   
   l = Len(HexValue)
   Val = 0
   For i = 1 To l
      p = InStr(1, HEXCHARS, Mid(HexValue, i, 1), vbTextCompare)
      If p = 0 Then
         Hex2Char = HexValue
         Exit Function
      End If
      
      Val = Val + (16 ^ (l - i)) * (p - 1)
   Next i
   
   Hex2Char = Chr(Val)

End Function

' FnTemplate: Plantilla RTF o HTML
' FnOut: Archivo de salida
' RepStrs(i, 0): nombre del campo
' RepStrs(i, 1): valor
' nRep: numero de datos en arreglo RepStrs
Public Function MergeFile(ByVal FnTemplate As String, ByVal FnOut As String, ByRef RepStrs() As String, ByVal nRep As Integer) As Boolean
   Dim iFd As Long, oFd As Long
   Dim Buf As String, ChrMask As String
   Dim i As Integer, l As Long
   
   MergeFile = False
   
   On Error Resume Next
      
   iFd = FreeFile()
   Open FnTemplate For Input As #iFd
   If ERR Then
      MsgErr FnTemplate
      Exit Function
   End If

   oFd = FreeFile()
   Open FnOut For Output As #oFd
   If ERR Then
      MsgErr FnOut
      Exit Function
   End If

   ChrMask = Left(RepStrs(0, 0), 1)
   l = 0

   Do Until EOF(iFd)
      Line Input #iFd, Buf
      l = l + 1
      
      For i = 0 To nRep
         If RepStrs(i, 0) <> "" Then
            If InStr(Buf, ChrMask) <> 0 Then
'               If InStr(Buf, "$ajuste$") And RepStrs(i, 0) = "$ajuste$" Then
'               Beep
'               End If
               Buf = ReplaceStr(Buf, RepStrs(i, 0), RepStrs(i, 1))
            Else
               Exit For
            End If
         End If
      Next i
      
      Print #oFd, Buf

   Loop

   Close #iFd
   Close #oFd

   If l < 5 Then
      Call AddLog("La plantilla '" & FnTemplate & "' sólo tiene " & l & " líneas.")
   End If

   MergeFile = True

End Function

' Si Fit = 0, la imagen crece hasta su tamaño
Public Function FitPicture(Pic As PictureBox, ByVal fname As String, ByVal wMax As Single, ByVal hMax As Single, Optional ByVal Fit As Boolean = 1, Optional bErrMsg As Boolean = 1) As Integer
   Dim wImg As Single, hImg As Single
   Dim wFact As Single, hFact As Single
   Dim sErr As String

   On Error Resume Next

   FitPicture = 0

   ' reponemos el tamaño original del picture
   Pic.visible = False
   If fname <> "" Then
      Set Pic.Picture = LoadPicture()
      
      If Fit Then
         Pic.Height = wMax
         Pic.Width = hMax
      
      End If
      
      Pic.AutoSize = Fit
      Pic.AutoRedraw = Not Fit
      
      Set Pic.Picture = LoadPicture(fname)
      If ERR Then
         FitPicture = ERR
         If bErrMsg Then
            sErr = ERR & ", " & Error
            Call AddLog("FitPicture: Error " & sErr & vbTab & fname)
            MsgBox1 "Error " & sErr & vbLf & fname, vbExclamation
         End If
         Exit Function
      End If
      wImg = Pic.Width
      hImg = Pic.Height
   Else
      wImg = Pic.Picture.Width / 1.75514096185738
      hImg = Pic.Picture.Height / 1.75077303648732
   End If
   
   'If Fit = 0 And wImg < wMax And hImg < hMax Then
   '   Pic.Visible = True
   '   Exit Function
   'End If
         
   If Fit = False Then
      Pic.visible = True
      Exit Function
   End If
   
   If wImg = wMax And hImg = hMax Then
      Pic.visible = True
      Exit Function
   End If
   
   If wImg > 0 Then
      wFact = wMax / wImg
   Else
      wFact = 1
   End If
   
   If hImg > 0 Then
      hFact = hMax / hImg
   Else
      hFact = 1
   End If
   
   Pic.AutoSize = False
   Pic.AutoRedraw = True
   
   If wFact > hFact Then
      Pic.Height = hMax
      Pic.Width = wImg * hFact
   Else
      Pic.Width = wMax
      Pic.Height = hImg * wFact
   End If
                  
   Pic.PaintPicture Pic, 0, 0, Pic.Width, Pic.Height, , , , , vbSrcCopy
   Pic.visible = True

End Function

Public Function ChkSystem(ByVal bMsg As Boolean) As Boolean
   Dim sDec As String, sMonDec As String
   Dim sThous As String, sMonThous As String
   Dim sDefDec As String, sDefThous As String
   Dim Msg As String
   
   ChkSystem = True
   
   If LCase(Format(DateSerial(2000, 1, 1), "mmm")) = "jan" Then ' inglés ?
      sDefDec = "."
      sDefThous = ","
   Else
      sDefDec = ","
      sDefThous = "."
   End If
   
   sDec = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal", sDefDec)
   sMonDec = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep", sDefDec)
   sThous = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sThousand", sDefThous)
   sMonThous = QryRegValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonThousandSep", sDefThous)

   Msg = "¡ ATENCIÓN !" & vbLf & "Los signos de decimales o separadores de miles no son consistentes en Números y Monedas." & vbLf & "Revise en Configuración Regional en el Panel de Control."
   
   If sDec <> "" And sMonDec <> "" And sDec <> sMonDec Then
      ChkSystem = False
      Call AddLog(Msg & " Num=" & Format(12345.64, DBLFMT2) & " Mon=" & Format(12345.64, "$" & DBLFMT2))
      
      If bMsg Then
         MsgBox1 Msg, vbCritical
      End If
      Exit Function
   End If
   
   If sThous <> "" And sMonThous <> "" And sThous <> sMonThous Then
      ChkSystem = False
      Call AddLog(Msg & " Num=" & Format(23456.74, DBLFMT2) & " Mon=" & Format(12345.64, "$" & DBLFMT2))
      
      If bMsg Then
         MsgBox1 Msg, vbCritical
      End If
      Exit Function
   End If
   
   If sDec <> "" And sThous <> "" And sDec = sThous Then
      ChkSystem = False
      Call AddLog(Msg & " Num=" & Format(34567.84, DBLFMT2) & " Mon=" & Format(12345.64, "$" & DBLFMT2))
      
      If bMsg Then
         MsgBox1 Msg, vbCritical
      End If
      Exit Function
   End If
   
End Function

Public Function Grep(ByVal fname As String, ByVal What As String, Optional ByVal From As Long = 1) As Long
   Dim Fd As Long, i As Long
   Dim Buf As String
   
   On Error Resume Next
   
   Fd = FreeFile()
   Open fname For Input As #Fd

   If From > 0 Then
      Seek #Fd, From
   End If

   Do Until EOF(Fd)
      
      Line Input #Fd, Buf

      i = InStr(1, Buf, What, vbTextCompare)
      If i Then
         Grep = From + i
         Exit Do
      End If
      
      From = From + Len(Buf) + 1
   
   Loop

   Close #Fd

End Function

Public Function GetAppVersion(Version As String, Fecha As String) As Long
   
   Version = W.Version
   
'   Fecha = Format(W.FVersion, "mmm d, yyyy")
   Fecha = IIf(W.FVersion > 1000, Format(W.FVersion, "d mmm yyyy"), "")  ' 24 jun 2019
   GetAppVersion = W.FVersion

End Function

' Normaliza una version de un programa del tipo v.r.d a vvv.rrr.dddd
Public Function NormVer(ByVal Version As String) As String
   Dim i As Integer, j As Integer
   Dim Buf As String
   
   i = InStr(Version, ".")
   If i = 0 Then
      Exit Function
   End If
   
   Buf = Right("000" & Left(Version, i - 1), 3)
   
   j = InStr(i + 1, Version, ".", vbBinaryCompare)
   Buf = Buf & "." & Right("000" & Mid(Version, i + 1, j - i - 1), 3)
   
   Buf = Buf & "." & Right("000" & Mid(Version, j + 1), 3)

   NormVer = Buf
   
End Function


' Transforma una IP 23.234.123.54 a 1A782323
Function IP2Hex(ByVal IP As String) As String
   Dim i As Integer, j As Integer
   Dim dIP As String, hIP As String

   IP = Trim(IP)

   hIP = ""
   For j = 1 To 4
      i = InStr(IP, ".")

      If i Then
         dIP = Left(IP, i) + 0
      Else
         dIP = IP + 0
      End If

      hIP = hIP & Right("00" & Hex(dIP), 2)
      IP = Mid(IP, i + 1)
   Next
         
   IP2Hex = hIP

End Function

Public Sub PamRandomize()
   Dim Tm As Double
   
   If lRndInit Then
      Exit Sub
   End If
   
   Tm = (Now * 10000)
   Tm = Tm - Int(Tm)

   Randomize Tm * 1000000
   lRndInit = True

End Sub

' Convierte un nombre de archivo a un nombre válido.
' No debe tener un path
Public Function ValidFName(ByVal Filename As String) As String
   Dim Chrs As String, i As Integer
   
   Chrs = "/\:*?""<>|"

   For i = 1 To Len(Chrs)
      Filename = Replace(Filename, Mid(Chrs, i, 1), "")
   Next i
   
   ValidFName = Trim(Filename)
   
End Function
' Imprime el contenido de un form usando las mismas posiciones de los objetos
' sólo maneja: Label, TextBox, Line, Frame, ComboBox, CheckBox, OptionButton y FlexGrid
Public Function PrtForm(Frm As Form, Optional ByVal bEndDoc As Integer = 1, Optional ByVal TopMargin As Integer = 0, Optional ByVal LeftMargin As Integer = 0, Optional Prt As Object = Nothing)
   Dim i As Integer, bCtrl As Boolean, Align As Integer, T As Integer
   Dim TName As String, Text As String
   Dim Ctrl As Control, CtrlP As Object
   Dim dx As Long, dy As Long, tdx As Long, Visib As Boolean, Aux As Integer
   Dim Bottom As Long, TagFirst As String, TBorder As Byte
   Dim dwX As Single, bWrap As Boolean, bDrawBorder As Boolean
   
   On Error Resume Next
   
   If Prt Is Nothing Then
      Set Prt = Printer
   End If
   
   Bottom = 0
   TagFirst = "-" ' los objetos con este TAG se imprimen primero
   For T = 0 To 1
   
      For i = 0 To Frm.Controls.Count - 1
      
         Set Ctrl = Frm.Controls(i)
         Debug.Print "[" & Ctrl.Name & "]"
         If W.InDesign And StrComp(Ctrl.Name, "Im_logoD", vbTextCompare) = 0 Then
            Beep
         End If
         
         If Left(Ctrl.Tag, 1) = "h" Then   ' hidden, no print
            GoTo Next_i
         ElseIf Left(Ctrl.Tag, 1) <> TagFirst Then
            GoTo Next_i
         End If
         
         TBorder = 0
         
         TName = LCase(TypeName(Ctrl))
         Text = ""
         dx = 0
         tdx = 0  ' para las checkbox o frames
         dwX = 0  ' para cuando el ancho del control es mayor que lo visible
         dy = 0
         bCtrl = True
         Visib = False
         Visib = Ctrl.visible
         Align = vbLeftJustify
         bWrap = False
         bDrawBorder = False
         
         If Visib = True Or Frm.visible = False Then
         
            Select Case TName
               Case "label":
                  If W.InDesign And StrComp(Ctrl.Name, "La_AnoMes", vbTextCompare) = 0 Then
                     Beep
                  End If
                  
                  Text = Ctrl.Caption
                  Align = Ctrl.Alignment
                  bWrap = (Ctrl.AutoSize = False)  '  And Align = vbLeftJustify)
               
               Case "frame":
                  Text = Ctrl.Caption
                  tdx = tdx + (Prt.TwipsPerPixelX * 50)
                  bDrawBorder = (Ctrl.BorderStyle <> 0)
               
               Case "textbox":
                  Debug.Print Ctrl.Name
                  Text = Ctrl.Text
                  TBorder = 1
                  Align = Ctrl.Alignment
                  bWrap = Ctrl.MultiLine
                                 
               Case "combobox":
                  TBorder = 2
                  If Ctrl.Style = vbComboDropdownList And Ctrl.ListIndex >= 0 Then
                     Text = Ctrl
                  Else
                     Text = Ctrl.Text
                  End If
               
               Case "checkbox", "optionbutton":
                  TBorder = 2
                  Text = Ctrl.Caption
               
               Case "picturebox", "graph":
                  TBorder = Ctrl.BorderStyle
                              
               Case "msflexgrid":
               Case "fedgrid", "fed2grid", "fed3grid", "fed4grid":  ' OJO: abajo hay que poner la misma lista
               Case "line":
               
               Case Else
                  bCtrl = False
                  Align = Ctrl.Alignment
                                 
            End Select
               
            If bCtrl Then
               Set CtrlP = Ctrl.Container
               Do While CtrlP.hWnd <> Frm.hWnd
                  If (CtrlP.visible = False And Frm.visible = True) Or CtrlP.Tag = "h" Then
                     GoTo Next_i
                  End If
                  
                  dx = dx + CtrlP.Left
                  dy = dy + CtrlP.Top
                  
                  Set CtrlP = CtrlP.Container
               Loop
            
               If Trim(Text) <> "" Then
                  
                  Prt.Font = Ctrl.Font
                  Prt.FontSize = Ctrl.FontSize
                  Prt.FontBold = Ctrl.FontBold
                  Prt.FontItalic = Ctrl.FontItalic
                  Prt.ForeColor = vbBlack
                  Prt.ForeColor = Ctrl.ForeColor
                  
                  If TBorder = 1 Then ' texbox
                     dy = dy + (Prt.TwipsPerPixelY * 2)
                     dx = dx + (Prt.TwipsPerPixelX * 23)   ' desplazamos un poco a la derecha
                     dwX = (Prt.TwipsPerPixelX * 23)       ' el ancho disminuye
                  ElseIf TBorder = 2 Then ' combobox y listbox
                     dy = dy + (Ctrl.Height - Prt.TextHeight(Text)) / 2
                     dx = dx + Prt.TextWidth(" ") / 2
                  End If
                  
                  'Align = Ctrl.Alignment
                  
                  If InStr(",checkbox,optionbutton,", "," & TName & ",") Then
                  
                     If Align = vbRightJustify Then
                        Align = vbLeftJustify
                     Else
                        tdx = Prt.TextWidth("[X]x")
                     End If
                  End If
                     
                  If Align = vbCenter Then
                     Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX

                     If bWrap = False Then
                        Prt.CurrentX = Prt.CurrentX + (Ctrl.Width - Prt.TextWidth(Text)) / 2
                     End If
                  ElseIf Align = vbRightJustify Then
                     If bWrap = False Then
                        Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX + Ctrl.Width - Prt.TextWidth(Text)
                     Else
                        Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX
                     End If
                  Else
                     Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left
                  End If
                  
                  Prt.CurrentY = TopMargin + dy + Ctrl.Top
                  
                  If bWrap Then
                     Call PrtBuf(Align, Prt.CurrentX, Prt.CurrentX + Ctrl.Width, Text, Prt, Prt.CurrentY + Ctrl.Height)
                  Else
                     Prt.Print Text
                  End If
                               
               End If
               
               ' los que tienen gráfica
               Select Case TName
                  Case "line":
                     Prt.DrawWidth = Ctrl.BorderWidth
                     Prt.Line (LeftMargin + dx + Ctrl.X1, TopMargin + dy + Ctrl.Y1)-(LeftMargin + dx + Ctrl.x2, TopMargin + dy + Ctrl.Y2)
                                 
                  Case "frame":
                     Prt.DrawWidth = 2
                     dy = dy + 60
                     
                     If Text <> "" Then
                        Aux = Prt.TextWidth(Text) + 180
                        
                        If bDrawBorder Then
                           Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + 60, TopMargin + dy + Ctrl.Top)
                           Prt.Line (LeftMargin + dx + Aux + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top)
                        End If
                     ElseIf bDrawBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top)
                     End If
                     
                     If bDrawBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                        
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                        Prt.Line (LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                     End If
                     
                  Case "checkbox", "optionbutton":
                     If Ctrl.Value Then
                        Text = "[x]"
                     Else
                        Text = "[  ]"
                     End If
                     
                     Prt.FontItalic = False
                     
                     If Ctrl.Alignment = vbRightJustify Then
                        Prt.CurrentX = LeftMargin + dx + Ctrl.Left + Ctrl.Width - Prt.TextWidth(Text)
                     Else
                        Prt.CurrentX = LeftMargin + dx + Ctrl.Left
                     End If
                  
                     Prt.CurrentY = TopMargin + dy + Ctrl.Top
                  
                     Prt.Print Text
                  
                  Case "picturebox", "graph":
                     Prt.CurrentX = LeftMargin + dx + Ctrl.Left
                     Prt.CurrentY = TopMargin + dy + Ctrl.Top
                  
                     Prt.PaintPicture Ctrl.Picture, Prt.CurrentX, Prt.CurrentY, Ctrl.Width, Ctrl.Height
                  
                     If TBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top)
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                        
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                        Prt.Line (LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height)
                     End If
                  
                  Case "msflexgrid":
                     Call FGrPrint(Ctrl, TopMargin, LeftMargin)
                  
                  Case "fedgrid", "fed2grid", "fed3grid", "fed4grid": ' Ojo: arriba debe estar la misma lista
                     Call FGrPrint(Ctrl, TopMargin, LeftMargin)
                  
               End Select
            End If
            
            If Prt.CurrentY > Bottom Then
               Bottom = Prt.CurrentY
            End If
            
         End If
Next_i:
      Next i
      
      TagFirst = ""
      
   Next T
   Prt.CurrentY = Bottom
   
   If bEndDoc Then
      Prt.EndDoc
   End If
   
End Function
' Imprime el contenido de un form usando las mismas posiciones de los objetos
' sólo maneja: Label, TextBox, Line, Frame, ComboBox, CheckBox, OptionButton y FlexGrid
' Utiliza el Tag de cada objeto para saber el orden de impresión, luego conviene usar frames
' y ponerle el mismo valor en el tag al frame y a sus objetos
Public Function PrtForm2(Frm As Form, Optional ByVal bEndDoc As Integer = 1, Optional ByVal TopMargin As Long = 0, Optional ByVal LeftMargin As Integer = 0, Optional Prt As Object = Nothing, Optional ByVal iTag As Integer = -1, Optional ByVal oTag As Integer = -1) As Long
   Dim i As Integer, bCtrl As Boolean, Align As Integer, T As Integer
   Dim TName As String, Text As String
   Dim Ctrl As Control, CtrlP As Object
   Dim dx As Long, dy As Long, tdx As Long, Visib As Boolean, Aux As Integer, dyPag As Long
   Dim Bottom As Long, TBorder As Byte
   Dim dwX As Single, bWrap As Boolean, bDrawBorder As Boolean
   
   On Error Resume Next
   
   If Prt Is Nothing Then
      Set Prt = Printer
   End If
   
   Bottom = 0
   dyPag = 0
   For T = iTag To oTag
   
      For i = 0 To Frm.Controls.Count - 1
      
         Set Ctrl = Frm.Controls(i)
         If Left(Ctrl.Tag, 1) = "h" Then   ' hidden, no print
            GoTo Next_i
         ElseIf T <> -1 And Val(Ctrl.Tag) <> T Then
            GoTo Next_i
         End If
         
         TBorder = 0
         
         TName = LCase(TypeName(Ctrl))
         Text = ""
         dx = 0
         tdx = 0  ' para las checkbox o frames
         dwX = 0  ' para cuando el ancho del control es mayor que lo visible
         dy = 0
         bCtrl = True
         Visib = False
         Visib = Ctrl.visible
         Align = vbLeftJustify
         bWrap = False
         bDrawBorder = False
         
         If Visib = True Or Frm.visible = False Then
         
            Select Case TName
               Case "label":
                  Text = Ctrl.Caption
                  Align = Ctrl.Alignment
                  bWrap = (Ctrl.AutoSize = False) '  And Align = vbLeftJustify)
               
               Case "frame":
                  Text = Ctrl.Caption
                  tdx = tdx + (Prt.TwipsPerPixelX * 50)
                  bDrawBorder = (Ctrl.BorderStyle <> 0)
               
               Case "textbox":
'                  Debug.Print Ctrl.Name
                  Text = Ctrl.Text
                  TBorder = 1
                  Align = Ctrl.Alignment
                  bWrap = Ctrl.MultiLine
                                 
               Case "combobox":
                  TBorder = 2
                  If Ctrl.ListIndex >= 0 Then
                     Text = Ctrl
                  End If
               
               Case "checkbox", "optionbutton":
                  TBorder = 2
                  Text = Ctrl.Caption
               
               Case "picturebox", "graph":
                  TBorder = Ctrl.BorderStyle
                              
               Case "msflexgrid":
               Case "line":
               
               Case Else
                  bCtrl = False
                  Align = Ctrl.Alignment
                                 
            End Select
               
            If bCtrl Then
               Set CtrlP = Ctrl.Container
               Do While CtrlP.hWnd <> Frm.hWnd
                  If (CtrlP.visible = False And Frm.visible = True) Or CtrlP.Tag = "h" Then
                     GoTo Next_i
                  End If
                  
                  dx = dx + CtrlP.Left
                  dy = dy + CtrlP.Top
                  
                  Set CtrlP = CtrlP.Container
               Loop
            
               If Trim(Text) <> "" Then
                  
                  Prt.Font = Ctrl.Font
                  Prt.FontSize = Ctrl.FontSize
                  Prt.FontBold = Ctrl.FontBold
                  Prt.FontItalic = Ctrl.FontItalic
                  Prt.ForeColor = vbBlack
                  Prt.ForeColor = Ctrl.ForeColor
                  
                  If TBorder = 1 Then ' texbox
                     dy = dy + (Prt.TwipsPerPixelY * 2)
                     dx = dx + (Prt.TwipsPerPixelX * 23)   ' desplazamos un poco a la derecha
                     dwX = (Prt.TwipsPerPixelX * 23)       ' el ancho disminuye
                  ElseIf TBorder = 2 Then ' combobox y listbox
                     dy = dy + (Ctrl.Height - Prt.TextHeight(Text)) / 2
                     dx = dx + Prt.TextWidth(" ") / 2
                  End If
                  
                  'Align = Ctrl.Alignment
                  
                  If InStr(",checkbox,optionbutton,", "," & TName & ",") Then
                  
                     If Align = vbRightJustify Then
                        Align = vbLeftJustify
                     Else
                        tdx = Prt.TextWidth("[X]x")
                     End If
                  End If
                     
                  If Align = vbCenter Then
                     Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX + (Ctrl.Width - Prt.TextWidth(Text)) / 2
                  ElseIf Align = vbRightJustify Then
                     If bWrap = False Then
                        Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX + Ctrl.Width - Prt.TextWidth(Text)
                     Else
                        Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left - dwX
                     End If
                  Else
                     Prt.CurrentX = LeftMargin + dx + tdx + Ctrl.Left
                  End If
                  
                  Prt.CurrentY = TopMargin + dy + Ctrl.Top + dyPag
                  
                  If bWrap Then
                     Call PrtBuf(Align, Prt.CurrentX, Prt.CurrentX + Ctrl.Width, Text, Prt, Prt.CurrentY + Ctrl.Height)
                  Else
                     Prt.Print Text
                  End If
                               
               End If
               
               ' los que tienen gráfica
               Select Case TName
                  Case "line":
                     Prt.DrawWidth = Ctrl.BorderWidth
                     Prt.Line (LeftMargin + dx + Ctrl.X1, TopMargin + dy + Ctrl.Y1 + dyPag)-(LeftMargin + dx + Ctrl.x2, TopMargin + dy + Ctrl.Y2 + dyPag)
                                 
                  Case "frame":
                     Prt.DrawWidth = 2
                     
'                     If Prt.CurrentY + Ctrl.Height + 120 > Prt.Height - TopMargin * 4 Then
                     If TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag > Prt.Height - TopMargin * 5 Then
                     
'                        dyPag = -(Bottom - TopMargin)
                        Prt.NewPage
                        Prt.CurrentY = 0
                        
                        Aux = PrtForm2(Frm, 0, TopMargin, LeftMargin, , 0, 0) ' El frame de encabezado
                        dyPag = Aux - Ctrl.Top + 60
'                        dyPag = dyPag + Prt.CurrentY
                     End If
                     
                     dy = dy + 60
                     
                     If Text <> "" Then
                        Aux = Prt.TextWidth(Text) + 180
                        
                        If bDrawBorder Then
                           Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + 60, TopMargin + dy + Ctrl.Top + dyPag)
                           Prt.Line (LeftMargin + dx + Aux + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + dyPag)
                        End If
                     ElseIf bDrawBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + dyPag)
                     End If
                     
                     If bDrawBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                        
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                        Prt.Line (LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                     End If
                     
                  Case "checkbox", "optionbutton":
                     If Ctrl.Value Then
                        Text = "[x]"
                     Else
                        Text = "[  ]"
                     End If
                     
                     Prt.FontItalic = False
                     
                     If Ctrl.Alignment = vbRightJustify Then
                        Prt.CurrentX = LeftMargin + dx + Ctrl.Left + Ctrl.Width - Prt.TextWidth(Text)
                     Else
                        Prt.CurrentX = LeftMargin + dx + Ctrl.Left
                     End If
                  
                     Prt.CurrentY = TopMargin + dy + Ctrl.Top + dyPag
                  
                     Prt.Print Text
                  
                  Case "picturebox", "graph":
                     Prt.CurrentX = LeftMargin + dx + Ctrl.Left
                     Prt.CurrentY = TopMargin + dy + Ctrl.Top + dyPag
                  
                     Prt.PaintPicture Ctrl.Picture, Prt.CurrentX, Prt.CurrentY, Ctrl.Width, Ctrl.Height
                  
                     If TBorder Then
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + dyPag)
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                        
                        Prt.Line (LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                        Prt.Line (LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + dyPag)-(LeftMargin + dx + Ctrl.Left + Ctrl.Width, TopMargin + dy + Ctrl.Top + Ctrl.Height + dyPag)
                     End If
                  
                  Case "msflexgrid":
                     Call FGrPrint(Ctrl, TopMargin, LeftMargin)
                  
               End Select
            End If
            
            If Prt.CurrentY > Bottom Then
               Bottom = Prt.CurrentY
            End If
            
'            If Prt.CurrentY > Prt.Height * 0.9 Then
'               dyPag = -(Bottom - TopMargin)
'               Prt.CurrentY = 0
'               Prt.NewPage
'
'               Call PrtForm2(Frm, 0, TopMargin, LeftMargin, , 0, 0)
'               dyPag = dyPag + Prt.CurrentY
'            End If
            
         End If
Next_i:
      Next i
            
   Next T
   Prt.CurrentY = Bottom
   
   If bEndDoc Then
      Prt.EndDoc
   End If
   
   PrtForm2 = Bottom
   
End Function


' Imprime el contenido de una grilla de tipo MsFlexGrid
Public Function FGrPrint(Gr As Control, Optional ByVal TopMargin As Integer = 0, Optional ByVal LeftMargin As Integer = 0)
   Dim i As Integer, bCtrl As Boolean, Align As Integer, T As Integer, Alig As Integer
   Dim Text As String
   Dim CtrlP As Object, GrParent As Object
   Dim dx As Long, dy As Long, yPage As Long, yBorder As Long, iPage As Integer, yBottom As Long, yAux1 As Long, yAux2 As Long
   Dim r As Integer, c As Integer, CWid As Integer, bDrawVert As Boolean
   
   On Error Resume Next
   
   Set CtrlP = Gr.Container
   Set GrParent = Gr.Parent
   
   yPage = 0
   iPage = 0
   yBorder = Printer.TextHeight("W") * 3
   dx = Gr.Left
   dy = Gr.Top
   bDrawVert = True
   
   If LCase(TypeName(Gr)) = "fedgrid" Then
      Set Gr = Gr.FlxGrid
   End If
            
   Do While CtrlP.hWnd <> GrParent.hWnd
      dx = dx + CtrlP.Left
      dy = dy + CtrlP.Top
      
      Set CtrlP = CtrlP.Container
   Loop
   

   Printer.DrawWidth = 2
   iPage = 1
   
   ' horizontales
   Printer.Line (LeftMargin + dx, TopMargin + dy)-(LeftMargin + dx + Gr.Width, TopMargin + dy)
   Printer.Line (LeftMargin + dx, TopMargin + dy + Gr.Height)-(LeftMargin + dx + Gr.Width, TopMargin + dy + Gr.Height)
   
   For r = 0 To Gr.rows - 1
   
      If bDrawVert Then
         bDrawVert = False
         
         If iPage = 1 Then
            If TopMargin + dy + Gr.Height < Printer.Height - yBorder Then
               yBottom = TopMargin + dy + Gr.Height
            Else
               yBottom = Printer.Height - yBorder
            End If
         Else
            If yBorder + Gr.Height - Gr.RowPos(r) < Printer.Height - yBorder Then
               yBottom = yBorder + Gr.Height - Gr.RowPos(r)
            Else
               yBottom = Printer.Height - yBorder
            End If
         End If
            
         
         'verticales
''         Printer.Line (LeftMargin + dx, TopMargin + dy)-(LeftMargin + dx, TopMargin + dy + Gr.Height)
''         Printer.Line (LeftMargin + dx + Gr.Width, TopMargin + dy)-(LeftMargin + dx + Gr.Width, TopMargin + dy + Gr.Height)
'         Printer.Line (LeftMargin + dx, TopMargin + dy)-(LeftMargin + dx, yBottom)
'         Printer.Line (LeftMargin + dx + Gr.Width, TopMargin + dy)-(LeftMargin + dx + Gr.Width, yBottom)
'
'         For c = 0 To Gr.Cols - 1
'
'            If c = Gr.FixedCols And Gr.FixedCols <> 0 Then
'               c = Gr.LeftCol
'               Printer.DrawWidth = 2
'            Else
'               Printer.DrawWidth = 1
'            End If
'
'            ' linea vertical
''            If iPage = 1 Then
''               Printer.Line (LeftMargin + dx + Gr.ColPos(c), TopMargin + dy - yPage)-(LeftMargin + dx + Gr.ColPos(c), TopMargin + dy + Gr.Height - yPage)
''            Else
''               Printer.Line (LeftMargin + dx + Gr.ColPos(c), yBorder)-(LeftMargin + dx + Gr.ColPos(c), TopMargin + dy + Gr.Height - yPage)
''            End If
'            If iPage = 1 Then
'               Printer.Line (LeftMargin + dx + Gr.ColPos(c), TopMargin + dy - yPage)-(LeftMargin + dx + Gr.ColPos(c), yBottom)
'            Else
'               Printer.Line (LeftMargin + dx + Gr.ColPos(c), yBorder)-(LeftMargin + dx + Gr.ColPos(c), yBottom)
'            End If
'
'         Next c
      End If
   
      If Gr.RowHeight(r) <> 0 Then
   
         If Gr.RowPos(r) + Gr.RowHeight(r) > Gr.Height Then
            Exit For
         End If
            
         If r = Gr.FixedRows And Gr.FixedRows <> 0 Then
            r = Gr.TopRow
            Printer.DrawWidth = 2
         Else
            Printer.DrawWidth = 1
         End If
         
         ' linea horizontal antes de los datos
         yAux1 = TopMargin + dy + Gr.RowPos(r) - yPage
         Printer.Line (LeftMargin + dx, yAux1)-(LeftMargin + dx + Gr.Width, yAux1)
                  
         Gr.Row = r
   
         If Printer.CurrentY > yBottom Then
            Printer.NewPage
            iPage = iPage + 1
            yPage = TopMargin + dy + Gr.RowPos(r) - yBorder
            
            yAux1 = TopMargin + dy + Gr.RowPos(r) - yPage
            Printer.Line (LeftMargin + dx, yAux1)-(LeftMargin + dx + Gr.Width, yAux1)

            bDrawVert = True
         End If
   
         For c = 0 To Gr.Cols - 1
      
            If Gr.ColWidth(c) <> 0 Then
      
               If Gr.ColPos(c) + Gr.ColWidth(c) > Gr.Width Then
                  Exit For
               End If
                        
               If c = Gr.FixedCols And Gr.FixedCols <> 0 Then
                  c = Gr.LeftCol
'                  Printer.DrawWidth = 2
'                  Printer.Line (LeftMargin + dx + Gr.ColPos(c), TopMargin + dy)-(LeftMargin + dx + Gr.ColPos(c), TopMargin + dy + Gr.Height)
               End If
               
               Text = Gr.TextMatrix(r, c)
                     
               If Trim(Text) <> "" Then
                  
                  Gr.Col = c
                  
                  Text = " " & Text & " "
                  
                  Printer.FontSize = Gr.CellFontSize
                  Printer.FontBold = Gr.CellFontBold
                  Printer.FontItalic = Gr.CellFontItalic
                  Printer.ForeColor = vbBlack
                  Printer.ForeColor = Gr.CellForeColor
                  
                  If r < Gr.FixedRows Then
                     Align = Gr.FixedAlignment(c)
                  Else
                     Align = Gr.ColAlignment(c)
                  End If
                             
                  CWid = Gr.ColWidth(c) - 10
                             
                  If Align = 4 Then ' flexAlignCenterCenter Then
                     Alig = vbCenter
                     Printer.CurrentX = LeftMargin + dx + Gr.ColPos(c) + (CWid - Printer.TextWidth(Text)) / 2
                  ElseIf Align = 7 Then ' flexAlignRightCenter Then
                     Alig = vbAlignRight
                     Printer.CurrentX = LeftMargin + dx + Gr.ColPos(c) + CWid - Printer.TextWidth(Text)
                  Else
                     Alig = vbAlignLeft
                     Printer.CurrentX = LeftMargin + dx + Gr.ColPos(c)
                  End If
                  
                  Printer.CurrentX = Printer.CurrentX
                  Printer.CurrentY = TopMargin + dy + Gr.RowPos(r) - yPage
                  
                  Do While Printer.TextWidth(Text) > CWid
                     Text = Left(Text, Len(Text) - 1)
                  Loop
                  
                  Printer.Print Text
                  'Call PrtLine(Alig, Printer.CurrentX, Printer.CurrentX + CWid, Text, True)
             
               Else
                  Printer.CurrentY = TopMargin + dy + Gr.RowPos(r) - yPage
                  Printer.Print ""
             
               End If
               
            End If
            
         Next c
         
         yAux2 = Printer.CurrentY
         Printer.Line (LeftMargin + dx, yAux1)-(LeftMargin + dx, yAux2)
         Printer.Line (LeftMargin + dx + Gr.Width, yAux1)-(LeftMargin + dx + Gr.Width, yAux2 + 100)
         
         For c = 0 To Gr.Cols - 1
      
            If c = Gr.FixedCols And Gr.FixedCols <> 0 Then
               c = Gr.LeftCol
               Printer.DrawWidth = 2
            Else
               Printer.DrawWidth = 1
            End If
            
            ' linea vertical
'            If iPage = 1 Then
'               Printer.Line (LeftMargin + dx + Gr.ColPos(c), TopMargin + dy - yPage)-(LeftMargin + dx + Gr.ColPos(c), TopMargin + dy + Gr.Height - yPage)
'            Else
'               Printer.Line (LeftMargin + dx + Gr.ColPos(c), yBorder)-(LeftMargin + dx + Gr.ColPos(c), TopMargin + dy + Gr.Height - yPage)
'            End If
            If iPage = 1 Then
               Printer.Line (LeftMargin + dx + Gr.ColPos(c), TopMargin + dy - yPage)-(LeftMargin + dx + Gr.ColPos(c), yAux2)
            Else
               Printer.Line (LeftMargin + dx + Gr.ColPos(c), yBorder)-(LeftMargin + dx + Gr.ColPos(c), yAux2)
            End If
            
         Next c
         Printer.CurrentY = yAux2
         
      End If ' Gr.RowHeight(r) <> 0

   Next r
   
   ' linea horizontal
   Printer.Line (LeftMargin + dx, TopMargin + dy + Gr.Height - yPage)-(LeftMargin + dx + Gr.Width, TopMargin + dy + Gr.Height - yPage)

   
End Function


Public Function ValidEmail(ByVal email As String) As Boolean
   Dim i As Integer, j As Integer, k As Integer, Buf As String, bValid As Boolean
   Dim ch As String * 1
   
   ValidEmail = False

   If Len(email) < 6 Then
      Exit Function
   End If

   k = InStr(email, ";")   ' mails separados por ;
   If k > 0 Then
      ' Si vienen varios mails, se valida uno a uno
      Do While k > 0
      
         Buf = Trim(Left(email, k - 1))
      
         bValid = ValidEmail(Buf)
         If bValid = False Then
            ValidEmail = False
            Exit Function
         End If
   
         email = Mid(email, k + 1)
         k = InStr(email, ";")
         
      Loop
   
   End If


   For i = 1 To Len(email)
      ch = LCase(Mid(email, i, 1))
      If Not ((ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or InStr("._@-", ch) <> 0) Then
         Exit Function
      End If
   Next i

   If InStr(email, "..") Or InStr(email, ".@") Or InStr(email, "@.") Then
      Exit Function
   End If

   i = InStr(email, "@")
   If i Then
      j = InStr(i, email, ".", vbBinaryCompare)
      If InStr(i + 1, email, "@", vbBinaryCompare) Then
         Exit Function
      End If
   End If
   
   If i = 0 Or j = 0 Then
      Exit Function
   End If
      
   i = rInStr(email, ".")
   j = Len(Mid(email, i))
   If j <> 3 And j <> 4 Then  ' con el .
      Exit Function
   End If
      
   ValidEmail = True
   
End Function
' para ver el fin de un archivo
Public Function UnixEoF(UnixFile As UnixFile_t) As Boolean

   If EOF(UnixFile.Fd) = False Then
      UnixEoF = False
   Else
      UnixEoF = (Len(UnixFile.Buf) <= 0)
   End If

End Function
' Lee archivos con registros que terminan en Lf y CrLf
' Los archivos deben abrirse FOR BINARY
' Se debe usar la función UnixEoF
Public Function UnixLineInput(UnixFile As UnixFile_t) As String
   Dim Buf As String
   Dim l As Long

   Do
      l = InStr(UnixFile.Buf, vbLf)
      If l Then
         Buf = Left(UnixFile.Buf, l - 1)
         If Right(Buf, 1) = vbCr Then
            UnixLineInput = Left(Buf, Len(Buf) - 1)
         Else
            UnixLineInput = Buf
         End If
         UnixFile.Buf = Mid(UnixFile.Buf, l + 1)
         Exit Do
      ElseIf EOF(UnixFile.Fd) Then
         UnixLineInput = UnixFile.Buf
         UnixFile.Buf = ""
         Exit Do
      End If
      
      If EOF(UnixFile.Fd) = False Then
         Buf = String(2048, " ")
         Get #UnixFile.Fd, , Buf
         Buf = Left(Buf, StrLen(Buf))
         UnixFile.Buf = UnixFile.Buf & Buf
      End If
   
   Loop

End Function

Public Function FindPrinter(ByVal DeviceName As String, Optional ByVal bSet As Boolean = 0) As Printer
   Dim Prt As Printer
   
   DeviceName = LCase(Trim(DeviceName))
   Set FindPrinter = Nothing
   
   For Each Prt In Printers
      If DeviceName = LCase(Trim(Prt.DeviceName)) Then
         Set FindPrinter = Prt
         
         If bSet Then
            Set Printer = Prt
         End If
      
         Exit For
      End If

   Next Prt

End Function

' Hace lo mismo que la InStr pero de atrás para adelante
' Ojo: el Start no pudo estar como primer parámetro, se puede poner ,,
Public Function rInStr(ByVal Where As String, ByVal What As String, Optional ByVal Start As Long = 0, Optional compare As VbCompareMethod = vbBinaryCompare) As Long
   Dim l As Long, L2 As Long
   
   L2 = Len(What)
   
   If Start <= 0 Then
      Start = Len(Where) - L2 + 1
   End If
   
   For l = Start To 1 Step -1
      If StrComp(Mid(Where, l, L2), What, compare) = 0 Then
         rInStr = l
         Exit Function
      End If
   Next l
   
   rInStr = 0

End Function

' Obtiene desde Command$ el parámetro indicado, se asume que viene /Param=....
Public Function GetCmdParam(ByVal Param As String, Optional ByVal Def As String = "", Optional ByVal Def2 As String = "") As String
   Dim Cmd As String, i As Integer, l As Integer
   
   If Def2 = "" Then
      Def2 = Def
   End If
   
   Cmd = Trim(Command)
   Param = "/" & Trim(Param)
   i = InStr(1, Cmd, Param & "=", vbTextCompare)
   If i = 0 Then
      i = InStr(1, Cmd, Param, vbTextCompare)
      If i = 0 Then
         GetCmdParam = Def
      Else
         l = i + Len(Param) - 1
         
         If l >= Len(Cmd) Or Mid(Cmd, l + 1, 1) = " " Then
            GetCmdParam = Def2 ' hay al menos un /param  pero no /param=aaa
         End If
      End If
   Else
      Cmd = Mid(Cmd, i + Len(Param) + 1)
      i = InStr(Cmd, " /") ' donde empieza el siguiente Param
      If i Then
         Cmd = Left(Cmd, i - 1)
      End If

      GetCmdParam = Trim(Cmd)
   End If

End Function



Public Function MakeRGB(ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte) As Long
   Dim RGB As Long

   MakeRGB = (red * 256# * 256#) + (green * 256#) + blue

End Function

' Convierte un path del tipo j:\aaa\bbb\ccc\..\ddd a j:\aaa\bbb\ddd
Public Function AbsPath(ByVal RelPath As String) As String
   Dim i As Long, Path As String
   Dim p As Integer
   
   Do
      i = InStr(1, RelPath, "\..\", vbBinaryCompare)
   
      If i = 0 Then
         Exit Do
      End If
      
      p = rInStr(RelPath, "\", i - 1, vbBinaryCompare)
      
      RelPath = Left(RelPath, p) & Mid(RelPath, i + 4)
      
   Loop

   AbsPath = RelPath

End Function

Public Function FCheckSum(ByVal fname As String) As Long
   Dim Fd As Long, Sum(3) As Byte, l As Long, Chk As Long
   Dim Buf As String
   Dim i As Integer, c As Integer
    
   On Error Resume Next
   
   If ExistFile(fname) = False Then
      FCheckSum = -1
      Exit Function
   End If
   
   Fd = FreeFile()
   Open fname For Binary Access Read As #Fd
   If ERR Then
      FCheckSum = -ERR
      Exit Function
   End If

   l = 0
   Do
      Buf = String(2048, " ")
      Get #Fd, , Buf
      
      Debug.Print Len(Buf)
      For i = 1 To Len(Buf)
         c = Asc(Mid(Buf, i, 1))
            
         Sum(l Mod 4) = Sum(l Mod 4) Xor c
         l = l + 1
      Next i

   Loop Until EOF(Fd)

   Close #Fd

   Chk = 0
   For i = 0 To 3
      Chk = Chk * 256 + Sum(i)
   Next i
   FCheckSum = Chk

End Function

Public Function GenCode2(ByVal Buf As String, Optional ByVal Seed As Long = 0) As String
   Dim Cod As String, d As Long
   Dim i As Long, c As Long, b As Long, ch As String
   Dim pb As Long, PC As Long, lc As Long, lb As Long, pa As Long

   If Buf = "" Then
      GenCode2 = ""
      Exit Function
   End If

   Seed = Seed + 5432
   Cod = "PLKANCHYTWK"
   lc = Len(Cod)
   lb = Len(Buf)
   b = CInt((Seed + 3 * lb) Mod 654) + lb
   
   d = (Asc("Z") - Asc("A")) + 1
   
   For i = 1 To (lb + lc) * ((Seed Mod 13) + 5)
      PC = ((i + b) Mod lc) + 1
      pb = ((i + b) Mod lb) + 1
      pa = ((PC + i) Mod lc) + 1
      c = Seed + i + PC + pb + pb * Asc(Mid(Cod, PC, 1)) + PC * Asc(Mid(Buf, pb, 1)) + i * Asc(Mid(Cod, PC, pa))
      ch = Chr(Asc("A") + c Mod d)
      Mid(Cod, PC, 1) = ch
      'Debug.Print pc & " - " & pb & " - " & c & " - " & Ch & " - " & Cod
   Next i

   c = 0
   For i = 1 To lc
      c = c + i * Asc(Mid(Cod, i, 1))
   Next i
   ch = Chr(Asc("A") + c Mod d)
   Cod = Cod & ch
   
   'Debug.Print pc & " - " & pb & " - " & c & " - " & Ch & " - " & Cod

   GenCode2 = Cod
   
End Function

Public Function GenCode(ByVal Buf As String, Optional ByVal Seed As Long = 0) As String
   Dim Cod As String, d As Long
   Dim i As Long, c As Long, b As Long, ch As String
   Dim pb As Long, PC As Long, lc As Long, lb As Long

   If Buf = "" Then
      GenCode = ""
      Exit Function
   End If

   Seed = Seed + 5432
   Cod = "PLKANCHYTWK"
   lc = Len(Cod)
   lb = Len(Buf)
   b = CInt(Seed Mod 654)
   
   d = (Asc("Z") - Asc("A")) + 1
   
   For i = 1 To (lb + lc) * ((Seed Mod 10) + 5)
      PC = ((i + b) Mod lc) + 1
      pb = ((i + b) Mod lb) + 1
      c = Seed + i + PC + pb + pb * Asc(Mid(Cod, PC, 1)) + PC * Asc(Mid(Buf, pb, 1))
      ch = Chr(Asc("A") + c Mod d)
      Mid(Cod, PC, 1) = ch
      'Debug.Print pc & " - " & pb & " - " & c & " - " & Ch & " - " & Cod
   Next i

   c = 0
   For i = 1 To lc
      c = c + i * Asc(Mid(Cod, i, 1))
   Next i
   ch = Chr(Asc("A") + c Mod d)
   Cod = Cod & ch
   
   'Debug.Print pc & " - " & pb & " - " & c & " - " & Ch & " - " & Cod

   GenCode = Cod
   
End Function

Public Function IsValidCode(ByVal Code As String) As Boolean
   Dim i As Integer, c As Integer, ch As String, d As Integer
   
   IsValidCode = False
   
   If Code = "" Or Len(Code) < 11 Then
      Exit Function
   End If
   
   d = (Asc("Z") - Asc("A")) + 1
   
   c = 0
   For i = 1 To 11
      c = c + i * Asc(Mid(Code, i, 1))
   Next i
   ch = Chr(Asc("A") + c Mod d)

   IsValidCode = (ch = Right(Code, 1))

End Function

Public Function CopyFile(ByVal Source As String, ByVal Dest As String, Optional bOverwrite As Boolean = 0, Optional bMsg As Boolean = 0) As Long
   Dim bCopy As Boolean, Rc As Long, Msg As String

   On Error Resume Next
   CopyFile = 0
   
   If bOverwrite Then
      bCopy = True
   ElseIf ExistFile(Dest) = False Then
      bCopy = True
   Else
      bCopy = False
   End If
   
   If bCopy Then
      ERR.Clear
      
      FileCopy Source, Dest
      
      If ERR = 0 Then
         CopyFile = -1
      Else
         Rc = ERR.Number
         Msg = ERR.Description
         CopyFile = ERR.Number
         Call AddLog("CopyFile: Error " & Rc & ", " & Msg & ". " & Source & " => " & Dest)
         
         If bMsg Then
            MsgBox1 "CopyFile: Error " & Rc & ", " & Msg & vbCrLf & "desde :" & Source & vbCrLf & "hacia: " & Dest, vbExclamation
         End If
      End If
   Else
      CopyFile = -2
   End If

End Function
#If K_SelPrinter Then
Public Function SelPrinter(Optional ByVal bMsg As Boolean = 0, Optional ByVal bUseCopies As Boolean = 0) As Long
   Dim Prt As Printer, DevName As String
   Dim W As Object
   Static oDevName As String
   
   If gPrtDlg Is Nothing Then
      Debug.Print "*** No se ha asignado la variable gPrtDlg ***"
      SelPrinter = -1
      Exit Function
   End If
   
   On Error Resume Next

   If oDevName = "" Then  ' el que tiene el sistema
      oDevName = VB.Printer.DeviceName
   End If
   
   SelPrinter = 0
   
   VB.Printer.EndDoc
   
   'gPrtDlg.flags = cdlPDPrintSetup ' Or cdlPDNoPageNums Or cdlPDNoSelection
   'gPrtDlg.flags = (cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDReturnIC Or cdlPDUseDevModeCopies)
   gPrtDlg.Flags = (cdlPDReturnDC Or cdlPDNoPageNums Or cdlPDNoSelection) ' Or cdlPDUseDevModeCopies)
   
   If bUseCopies Then
      gPrtDlg.Flags = gPrtDlg.Flags Or cdlPDUseDevModeCopies
   End If
   
   gPrtDlg.CancelError = True
   
   gPrtDlg.PrinterDefault = True ' ** No eliminar
   Printer.TrackDefault = True   ' ** No eliminar

   Call gPrtDlg.ShowPrinter

   If ERR <> 0 Then
   
      SelPrinter = ERR
   
      If ERR <> cdlCancel And bMsg = True Then
         MsgErr ""
      End If
   
   Else
      ' Todo este chamullo es para dejar la impresora por defecto que estaba
      Dim p As Printer
      
      DevName = VB.Printer.DeviceName
      If oDevName <> DevName Then
         Set W = CreateObject("WScript.Network")
         If Not W Is Nothing Then
            Call W.SetDefaultPrinter(oDevName) ' repone el default printer del equipo
            Set W = Nothing
         
            ' Ahora buscamos el seleccionado
            For Each p In VB.Printers
               If p.DeviceName = DevName Then
                  Set VB.Printer = p
                  Exit For
               End If
            Next p
         End If
      End If
     
   End If
      
   VB.Printer.Copies = gPrtDlg.Copies
    
   If gPrtDlg.Orientation = cdlLandscape Then  ' ccOrientationHorizontal Then
      Printer.Orientation = vbPRORLandscape
   Else
'      Printer.Orientation = vbPRORPortrait
      Printer.Orientation = gPrtDlg.Orientation

   End If
      
   Set gPrinter = VB.Printer
'   For Each Prt In Printers
'      Debug.Print Prt.DeviceName, Prt.hDC
'      If (Prt.hDC Mod 1000000) = (gPrtDlg.hDC Mod 1000000) Then
'         Set gPrinter = Prt
'         Exit For
'      End If
'   Next Prt
'
'   Call SetIniString(gIniFile, "Config", "Printer", Printer.DeviceName)

   DoEvents

   Debug.Print "Printer: " & Printer.DeviceName & vbTab & "Size: " & Printer.PaperSize & vbTab & Format(Now, "hh:nn:ss")
'   Debug.Print "gPrinter: " & gPrinter.DeviceName & vbTab & Format(Now, "hh:nn:ss")

End Function
#End If
' Transforma un número a Hexadecimal, pero de largo L
Public Function Hex2(ByVal Number As Long, Optional ByVal l As Byte = 0)
   Dim Buf As String

   Buf = Hex(Number)
   If l <> 0 And Len(Buf) < l Then
      Buf = Right(String(l, "0") & Buf, l)
   End If
   
   Hex2 = Buf
   
End Function
' Transforma un número a Hexadecimal, pero de largo L
Public Function Bin(ByVal Number As Long)
   Dim Buf As String

   Buf = ""
   Do Until Number = 0
      Buf = (Number Mod 2) & Buf
      Number = Number \ 2
   Loop

   Bin = Buf
End Function

' Normaliza un string del tipo "1.12.3.1" a "0001.0012.0003.0001"
' para que pueda compararse
Public Function NormVersion(ByVal Ver As String, Optional ByVal SecLen As Byte = 5) As String
   Dim i As Integer, j As Integer, Buf As String, Fill As String

   Fill = String(SecLen, "0")
   
   i = 1
   Buf = Ver
   Do
      j = InStr(i, Buf, ".", vbBinaryCompare)
      If j Then
         Buf = Left(Buf, i - 1) & Right(Fill & Mid(Buf, i, j - i), SecLen) & Mid(Buf, j)
         j = InStr(i, Buf, ".", vbBinaryCompare)
         i = j + 1
      Else
         Buf = Left(Buf, i - 1) & Right(Fill & Mid(Buf, i), SecLen)
         Exit Do
      End If
   Loop

   NormVersion = Buf

End Function


' Le pone ceros a la izq si es numerico
Public Function NormalizeCod(ByVal Cod As String, ByVal Ln As Integer) As String

   Cod = Trim(Cod)
   If IsNumeric(Cod) Then
      NormalizeCod = Right(String(Ln + 2, "0") & Cod, Ln)
   Else
      NormalizeCod = Cod
   End If

End Function

Public Function DelFile(ByVal fname As String, Optional ByVal bMsg As Boolean = 0, Optional ByVal bLog As Boolean = 1, Optional ByVal bErrNoExist As Boolean = 0) As Long
   Dim Desc As String

   On Error Resume Next
   Kill fname
   
   If ERR.Number = 53 And bErrNoExist = False Then ' no importa si no existe
      ERR.Number = 0
      ERR.Description = ""
   End If
   
   DelFile = ERR
   Desc = ERR.Description
   
   If ERR.Number Then
      If bLog Then
         Call AddLog("Error " & ERR.Number & ", " & Desc & " al borrar '" & fname & "'.")
      End If
      
      ERR.Number = DelFile
      ERR.Description = Desc
         
      If bMsg Then
         If ERR.Number = 70 Then
            Desc = "El podría estar abierto." & vbCrLf
         Else
            Desc = ""
         End If
      
         MsgErr Desc & fname
      End If
   End If
   
End Function
' Para usar en las chekbox
Public Function ValSiNo(ByVal SiNo As String, Optional Otro As Integer = VAL_OTRO) As Byte

   SiNo = UCase(Left(Trim(SiNo), 1))

   If SiNo = "N" Then
      ValSiNo = VAL_NO
   ElseIf SiNo = "S" Then
      ValSiNo = VAL_SI
   Else
      ValSiNo = Otro
   End If

End Function
Public Function ToggleSiNo(ByVal SiNo As String) As String

   SiNo = UCase(Left(Trim(SiNo), 1))

   If SiNo = "N" Then
      ToggleSiNo = gSiNo(VAL_SI)
   ElseIf SiNo = "S" Then
      ToggleSiNo = gSiNo(VAL_NO)
   Else
      ToggleSiNo = gSiNo(VAL_SI)
   End If

End Function

' Sirve para CheckBox y RadioButton
Function ChkSiNo(Ctrl As CheckBox) As String
   
   Select Case Ctrl.Value
      Case vbChecked:
         ChkSiNo = "'S'"
      Case vbUnchecked:
         ChkSiNo = "'N'"
      Case Else:
         ChkSiNo = " NULL "
      
   End Select
      
End Function
Public Function FmtSiNo(ByVal SiNo As Integer, Optional bOtro As Boolean = 1, Optional ByVal StrOtro = "") As String

   If SiNo = VAL_NO Then
      FmtSiNo = gSiNo(VAL_NO) ' No
   ElseIf SiNo = VAL_SI Or bOtro = False Then
      FmtSiNo = gSiNo(VAL_SI) ' Si
   ElseIf StrOtro = "" Then
      FmtSiNo = "?"
   Else
      FmtSiNo = StrOtro
   End If

End Function
' Encripta el texto según la llave
Public Function FwEncrypt1(ByVal Text As String, ByVal Key As Long) As String
   Dim oText As String, i As Integer, ch As String, x As Long, l As Integer, c As Integer

   oText = ""
   l = Len(Text)
   If l <= 0 Then
      Exit Function
   End If
   
   x = Key + l * 17
   For i = 1 To l
   
      x = x + i
      If i > 1 Then
         x = x + Asc(Mid(Text, i - 1, 1))
      End If
         
      c = (Asc(Mid(Text, i, 1)) + (x Mod 128)) Mod 256
   
      oText = Right("00" & Hex(c), 2) & oText
   Next i
   
   'Debug.Print "Encrypt: [" & oText & "]"
   FwEncrypt1 = oText

End Function
' Desencripta el texto según la llave
Public Function FwDecrypt1(ByVal CrText As String, ByVal Key As Long) As String
   Dim oText As String, i As Integer, ch As String, x As Long, l As Integer, c As Integer

   oText = ""
   l = Len(CrText)
   If l < 2 Or l Mod 2 <> 0 Then
      Exit Function
   End If
   
   l = l / 2
   
   x = Key + l * 17
   For i = l To 1 Step -1
            
      x = x + (l - i + 1)
      If i < l Then
         x = x + Asc(Mid(oText, l - i, 1))
      End If
               
      c = (Val("&H" & Mid(CrText, i * 2 - 1, 2)) + 256 - (x Mod 128)) Mod 256
   
      ch = Chr(c)

      oText = oText & ch
   Next i
   
   'Debug.Print "Decrypt: [" & oText & "]"
   FwDecrypt1 = oText
   
End Function
' Para guardar datos en el Cfg encriptados
Public Function GetCfgCrypt(ByVal CfgFile As String, ByVal Section As String, ByVal Key As String, ByVal kDef As String, ByVal Seed As Long) As String
   Dim Valor As String, kValor As String, i As Integer

   kValor = GetIniString(CfgFile, Section, "k" & Key)
   If kValor = "" Then
      Valor = GetIniString(CfgFile, Section, Key)
      If Valor = "" Then
         kValor = kDef
      Else
         Valor = Rnd(Now) & "#" & Space(5) & Valor & Space(13)
         kValor = FwEncrypt1(Valor, Seed)
'        Debug.Print Key, Trim(Valor), kValor
         
         Call SetIniString(CfgFile, Section, "k" & Key, kValor)
         Call SetIniString(CfgFile, Section, Key, vbNullString)
      End If
   End If
   
   Valor = FwDecrypt1(kValor, Seed)
   i = InStr(Valor, "#")
   If i > 0 Then
      Valor = Mid(Valor, i + 1)
   End If
   
'   Debug.Print Key, Trim(Valor), kValor

   GetCfgCrypt = Trim(Valor)

End Function
Public Function MCaption(Mn As Menu)
   Dim Txt As String

   Txt = ReplaceStr(Mn.Caption, "&", "")
   MCaption = ReplaceStr(Txt, "...", "")

End Function

' En OW la fecha viene como syyddd: Siglo-Año-días
' 106010: siglo 21, año 2006, 10 días del año
Public Function Julian2Date(ByVal JDate As Long) As Long
   Dim yy As Integer, ddd As Integer

   yy = 2000 + ((JDate \ 1000) - 100)

   ddd = JDate Mod 1000

   Julian2Date = DateSerial(yy, 1, ddd)

End Function

' En OW la fecha viene como syyddd: Siglo-Año-días
' 106010: siglo 21, año 2006, 10 días del año
Public Function Date2Julian(ByVal vbDate As Long) As Long
   Dim yy As Integer, ddd As Integer, s As Integer

   yy = Year(vbDate)
   
   s = (yy \ 100) - 19 ' siglo 21 ==> 1
   ddd = vbDate - DateSerial(yy, 1, 1) + 1

   Date2Julian = (s * 100000#) + ((yy Mod 100#) * 1000) + ddd

End Function

Public Function Min(ByVal Num1 As Double, ByVal Num2 As Double) As Double
   If Num1 < Num2 Then
      Min = Num1
   Else
      Min = Num2
   End If
End Function
Public Function Max(ByVal Num1 As Double, ByVal Num2 As Double) As Double
   If Num1 > Num2 Then
      Max = Num1
   Else
      Max = Num2
   End If
End Function
' Por el W Vista hay que poner el ini en otra parte
Public Function GetIniFile(ByVal AppName As String) As String
   Dim Path As String
   Dim myOS As OSVERSIONINFOEX_T, Ver As Single

   GetIniFile = AppName & ".ini"

   myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
   'try win2000 version
   If GetVersionEx(myOS) Then
      
      Ver = Val(myOS.dwMajorVersion & "." & myOS.dwMinorVersion)
      
      If Ver > 5.1 Then  ' si es posterior a XP sp 2
         On Error Resume Next
         Path = Left(W.WinDir, 2) & "\Fairware"
         MkDir Path
         
'         Path = Path & "\" & AppName  26 abr 2016
'         MkDir Path
         
         GetIniFile = Path & "\" & AppName & ".ini"
      End If
   End If
   
End Function

Public Function OpenFile(ByVal hWnd As Long, ByVal fname As String, Optional ByVal ShowMode As VbAppWinStyle = vbMaximizedFocus, Optional ByVal nSeg As Long = 5) As Long
   Dim Ext As ExtInfo_t, Rc As Long, bOpened As Boolean
   Dim i As Integer, Cmd As String, Msg As String
            
   On Error Resume Next
   
   bOpened = False
   Rc = -111
   
   If ExistFile(fname) = False Then
      MsgBox1 "Archivo no encontrado" & vbCrLf & fname, vbExclamation
      Exit Function
   End If
   
   If bOpened = False Then
      i = rInStr(fname, ".")
      If i Then
         If GetExtInfo(Mid(fname, i), Ext) Then ' encontró una aplicación que lo abre
            If Ext.OpenCmd <> "" Then
               Cmd = GenCmd(Ext, "open", fname)
               Rc = ExecCmd(Cmd, ShowMode, nSeg * 1000)
               If Rc = 0 Then
                  bOpened = True
               Else
                  Msg = "Error " & Rc & " al ejecutar" & vbCrLf & Cmd
               End If
            End If
         Else
            Msg = "No hay ninguna aplicación asociada a la extensión " & Mid(fname, i) & "."
         End If
      End If
   End If

   ' 22 may 2017: por algún motivo no funciona en Win 10, asi que lo ponemos despues
   If bOpened = False Then
      ' Si no encontró una app, probamos de otra forma
      DoEvents
      Rc = ShellExecute(hWnd, "open", fname, "", "", ShowMode)
      If Rc > 32 Then
         bOpened = True
         Rc = 0
      Else
         If Msg <> "" Then
            Msg = Msg & vbCrLf
         End If
         Msg = Msg & "Error " & Rc & " al abrir el archivo " & vbCrLf & fname
      End If
   End If

   If bOpened = False And Msg <> "" Then
      MsgBox1 Msg, vbExclamation
   End If

   OpenFile = Rc
   
End Function
' Valida la fortaleza de una clave
Public Function StrongPassw(ByVal Passw As String, ByVal MinLen As Integer, ByVal bNumbers As Boolean, ByVal bSymbols As Boolean, Msg As String) As Integer
   Dim bHas As Boolean, i As Integer, l As Integer
   Const Symbols  As String = ".,:-;+_$%&/|@\=(){#}*?"

   StrongPassw = 0
   
   If Len(Passw) < MinLen Then
      StrongPassw = 4
      Msg = "La clave debe tener al menos " & MinLen & " caracteres."
      Exit Function
   End If

   If (LCase(Passw) Like "[a-zá-ú]*[a-zá-ú]*") = False Then
      StrongPassw = 3
      Msg = "La clave debe tener al menos dos letras y comenzar con una."
      Exit Function
   End If

   If bNumbers Then
      If (LCase(Passw) Like "[a-zá-ú]*#*") = False Then
         StrongPassw = 2
         Msg = "La clave debe tener al menos un dígito."
         Exit Function
      End If
   End If

   If bSymbols Then
      If (LCase(Passw) Like "[a-zá-ú]*[" & Symbols & "]*") = False Then
         StrongPassw = 1
         Msg = "La clave debe tener al menos un símbolo entre [" & Symbols & "]."
         Exit Function
      End If
   End If

End Function

Public Function AppendLines(ByVal Buf1 As String, ByVal Buf2 As String) As String

   Buf1 = RTrimLF(Buf1)
   Buf2 = RTrimLF(Buf2)

   If Buf1 = "" Then
      AppendLines = Buf2
   ElseIf Buf2 = "" Then
      AppendLines = Buf1
   Else
      AppendLines = Buf1 & vbCrLf & Buf2
   End If

End Function

Public Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Load XML
    objDom.async = False
    objDom.loadXML XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    ' Send XML command
    objXmlHttp.Send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.ResponseText
    
    ' Close object
    Set objXmlHttp = Nothing
        
    ' Return result
    PostWebservice = strRet
    
Exit Function
Err_PW:
    PostWebservice = "Error: " & ERR.Number & " - " & ERR.Description

End Function


Public Function FwDateSerial(ByVal Ano As Integer, ByVal Mes As Integer, ByVal Dia As Integer) As Long

   If Ano <= 0 Or Mes <= 0 Or Dia <= 0 Then
      FwDateSerial = 0
   Else
      FwDateSerial = DateSerial(Ano, Mes, Dia)
   End If
   
End Function

Public Function GetTagValue(ByVal Buf As String, ByVal Tag As String, Optional ByVal iBuf As Integer = 1) As String
   Dim i As Integer, j As Integer

   i = InStr(iBuf, Buf, "<" & Tag & ">", vbTextCompare)
   If i > 0 Then
      i = i + 2 + Len(Tag)
      
      j = InStr(i + 2, Buf, "</" & Tag & ">", vbTextCompare)
      
      GetTagValue = Mid(Buf, i, j - i)
      
   End If
   
End Function
Public Function AddTag(ByVal Tag As String, ByVal Valor As String, Optional ByVal Atrib As String = "") As String
   Dim Buf As String
   
   Buf = "<" & Tag
   
   If Len(Atrib) > 2 Then
      Buf = Buf & " " & Atrib
   End If
   
   Buf = Buf & ">" & Valor & "</" & Tag & ">"

   AddTag = Buf
   
End Function

' Info( 34,1 ) as string
' Info( i, 0 ): nombre de la propiedad (autor, título, etc)
' Info( i, 1 ): valor de la propiedad
Private Function GetFileInfo(ByVal Path As String, ByVal fname As String, Info() As String) As Long
   Dim arrHeaders(0 To 34), objShell As Object, objFolder As Object
   Dim i As Integer, strFileName As Object

   Set objShell = CreateObject("Shell.Application")
   If objShell Is Nothing Then
      GetFileInfo = 1
      Exit Function
   End If
   
   Set objFolder = objShell.Namespace("" & Path) ' el "" & es neceario.. no se porqué
   If objFolder Is Nothing Then
      GetFileInfo = 2
      Exit Function
   End If

   For i = 0 To UBound(Info)
       Info(i, 0) = objFolder.GetDetailsOf(objFolder.Items, i)
   Next

   GetFileInfo = 3  ' Por si no lo encuentra

   For Each strFileName In objFolder.Items
      If StrComp(strFileName, fname, vbTextCompare) = 0 Then
'         Debug.Print
         For i = 0 To UBound(Info)
            Info(i, 1) = objFolder.GetDetailsOf(strFileName, i)
'            If Len(Info(i, 1)) > 0 Then
'               Debug.Print i, Info(i, 0); ":"; Info(i, 1)
'            End If
         Next
         GetFileInfo = 0
         Exit For
      End If
   Next

   Set objFolder = Nothing
   Set objShell = Nothing
   
End Function

' Para usarla en KeyPress de las combobox que no tienen MaxLength
Public Sub CbMaxLength(KeyAscii As Integer, ByVal Txt As String, ByVal MaxLen As Integer)

   If Len(Txt) >= MaxLen And KeySys(KeyAscii) = False Then
      KeyAscii = 0
   End If

End Sub

Public Function GetFileAuthor(ByVal FilePath As String, Author As String) As Long
   Dim objDoc As Object, Rc As Long
   
   Rc = 0
   On Error Resume Next
   
   Set objDoc = CreateObject("DSOFile.OleDocumentProperties")
   If objDoc Is Nothing Then
      Rc = Error.Number
      If Rc = 0 Then
         Rc = 33
      End If
      Call AddLog("Falta registrar la componente dsofile.dll")
      
      GetFileAuthor = Rc
      Exit Function
   End If

   Call objDoc.Open(FilePath)

'   GetFileSummary = objDoc.SummaryProperties.Comments
   Author = objDoc.SummaryProperties.Author

   objDoc.Close
   Set objDoc = Nothing

   GetFileAuthor = 0

End Function

Function ParseDir(ByVal Direccion As String, Tipo As String, Calle As String, Numero As String, CasaDpto As String) As String
   Dim Idx As Integer, IdxNum As Integer, IdxDpto As Integer, ch As String, IdxSep As Integer
   Dim i As Integer
   Dim Separadores As String
   
   Separadores = " ,.;-("
   
   Idx = 0
   IdxNum = 0
   IdxDpto = 0
   
   Calle = ""
   Numero = ""
   CasaDpto = ""

   i = InStr(1, Direccion, "Avenida ", vbTextCompare)
   If i > 0 Then
      Tipo = "Avenida"
      Direccion = ReplaceStr(Direccion, "Avenida ", "")
   Else
      i = InStr(1, Direccion, "Av.", vbTextCompare)
      If i > 0 Then
         Tipo = "Avenida"
         Direccion = Trim(ReplaceStr(Direccion, "Av.", ""))
      Else
         i = InStr(1, Direccion, "Av ", vbTextCompare)
         If i > 0 Then
            Tipo = "Avenida"
            Direccion = ReplaceStr(Direccion, "Av ", "")
         End If
      End If
   End If
      
   If Tipo = "" And InStr(1, Direccion, "Calle", vbTextCompare) Then
      Tipo = "Calle"
   End If
   
   For i = 0 To 9
      ch = i
      Idx = InStr(1, Direccion, ch, vbBinaryCompare)
      If Idx > 0 Then
         If Idx < IdxNum Or IdxNum = 0 Then
            IdxNum = Idx
         End If
      End If
   
   Next i
   
   If IdxNum = 0 Then
      Calle = Direccion
      ParseDir = Direccion
      
      Exit Function
          
   End If
   
   Calle = Left(Direccion, IdxNum - 1)
   Calle = Trim(ReplaceStr(Calle, "#", ""))
   
   IdxSep = 0
   For i = 1 To Len(Separadores)
   
      Idx = InStr(IdxNum + 1, Direccion, Mid(Separadores, i, 1))
      
      If Idx > 0 Then
         If Idx < IdxDpto Or IdxDpto = 0 Then
            IdxSep = Idx
         End If
      End If
      
   Next i
   
   If IdxSep = 0 Then
      Numero = Trim(Mid(Direccion, IdxNum))
      ParseDir = Calle & " " & Numero
      Exit Function
   End If
   
   If IdxSep > 0 Then
      Numero = Val(Mid(Direccion, IdxNum, IdxSep - IdxNum))
   End If
   
   For i = Asc("A") To Asc("Z")
      ch = Chr(i)
      Idx = InStr(IdxSep + 1, UCase(Direccion), ch, vbTextCompare)
      
      If Idx > 0 Then
         If Idx < IdxDpto Or IdxDpto = 0 Then
            IdxDpto = Idx
         End If
      End If
            
   Next i
   
   If IdxDpto > 0 Then
      CasaDpto = Trim(Mid(Direccion, IdxDpto))
   End If

   If IdxDpto = 0 Then

      For i = 0 To 9
         ch = str(i)
         IdxDpto = InStr(IdxSep + 1, Direccion, ch, vbBinaryCompare)
         
         If IdxDpto > 0 Then
            CasaDpto = Trim(Mid(Direccion, IdxDpto))
         End If
         
      Next i
   End If
   
   ParseDir = Calle & " " & Numero

End Function

Public Function ConvSysVars(ByVal Buf As String) As String
   Dim i As Integer, j As Integer, k As Integer, Var As String, Value As String

   i = 1
   Do
      j = InStr(i, Buf, "%", vbBinaryCompare)
      If j <= 0 Then
         Exit Do
      End If
      
      k = InStr(j + 1, Buf, "%", vbBinaryCompare)
      If k <= 0 Then
         Exit Do
      End If
     
      Var = Mid(Buf, j + 1, k - j - 1)
      Value = Environ(Var)
      
      If Value <> "" Then
         Buf = ReplaceStr(Buf, "%" & Var & "%", Value)
      End If
      
      i = j + 1
   Loop

   ConvSysVars = Buf
   
End Function

' Busca en Buf los LF y los transforma en CR-LF
Public Function Lf2CrLf(ByVal Buf As String) As String
   Dim i As Integer, j  As Integer

   j = 1
   Do
      i = InStr(j, Buf, vbLf, vbBinaryCompare)
      
      If i > 1 Then
         If Mid(Buf, i - 1, 1) <> vbCr Then
            Buf = Left(Buf, i - 1) & vbCr & Mid(Buf, i)
         End If
         j = i + 1
      Else
         Exit Do
      End If
   Loop

   Lf2CrLf = Buf
End Function

Public Function SetClipText(ByVal Texto As String) As Long

   On Error Resume Next

   Clipboard.Clear
   Clipboard.SetText Texto

   SetClipText = ERR.Number
   If ERR.Number Then
      MsgErr "Error al copiar al portapapeles."
   End If

End Function

Public Function FwFindXmlTag(ByVal xml As String, ByVal Tag As String, ByVal Indice As Long, ByVal iPos As Long, ByVal fPos As Long, Fin As Long) As Long
   Dim t0 As Long, t1 As Long, t2 As Long, t3 As Long, i As Integer, lTag As Integer, lFinTag As Integer
   Dim Fld As String, l As Integer
   
   i = InStr(Tag, "|")
   If i > 0 Then
      Fld = Trim(Mid(Tag, i + 1))
      Tag = Trim(Left(Tag, i - 1))
   Else
      Fld = ""
      Tag = Trim(Tag)
   End If
   
   lTag = Len(Tag)
   
   If iPos = 0 Then
      iPos = 1
   End If
   
   t0 = iPos
   i = 0
   Fin = 0
   Do
      ' buscamos el comienzo del tag, <tag> o <tag ....>
      t1 = InStr(t0, xml, "<" & Tag & ">", vbTextCompare)
      t2 = InStr(t0, xml, "<" & Tag & " ", vbTextCompare)
      If t2 > 0 And (t2 < t1 Or t1 = 0) Then
         t1 = t2
      End If
      
      If t1 <= 0 Or (fPos > 0 And t1 > fPos) Then
         FwFindXmlTag = -1  ' No existe en el rango
         Exit Function
      End If

      If Fld = "" Then  ' por si es <tag fld='...'>
         t2 = InStr(t1 + lTag + 2, xml, "/>", vbTextCompare)
         t3 = InStr(t1 + lTag + 2, xml, "</" & Tag & ">", vbTextCompare)
         
         If fPos > 0 And t2 > fPos Then
            t2 = 0
         End If
         
         If fPos > 0 And t3 > fPos Then
            t3 = 0
         End If
         
         If t2 = 0 And t3 = 0 Then
            FwFindXmlTag = -2  ' No existe el fin en el rango
            Exit Function
         End If
         
         If t3 > 0 Then   ' </tag
            t2 = t3
            t1 = InStr(t1 + lTag, xml, ">", vbBinaryCompare) + 1
            lFinTag = lTag + 3
         
         Else ' If t2 > 0 And (t2 < t3 Or t3 = 0) Then  ' <tag  ...... />
            t2 = t1
            lFinTag = 0
'         Else
'            t2 = t3
'            t1 = InStr(t1 + lTag, xml, ">", vbBinaryCompare) + 1
'            lFinTag = lTag + 3
         End If
      Else ' Fld <> ""
         l = 2
         t2 = InStr(t1 + lTag + 1, xml, "/>", vbBinaryCompare)
         t3 = InStr(t1 + lTag + 1, xml, ">", vbBinaryCompare)
         
         If fPos > 0 And t2 > fPos Then
            t2 = 0
         End If
         
         If fPos > 0 And t3 > fPos Then
            t3 = 0
         End If
         
         If t3 > 0 And (t3 < t2 Or t2 = 0) Then
            l = 1
            t2 = t3
         End If
                  
         If t2 = 0 Then
            FwFindXmlTag = -2  ' No existe el campo en el rango
            Exit Function
         End If
         
         If t2 > 0 Then
            t1 = t1 + lTag + l
            lFinTag = l
         End If
      
      End If
         
      i = i + 1
      
      If i = Indice Or Indice = 0 Then
         Fin = t2
         FwFindXmlTag = t1
         Exit Function
      End If
         
      t0 = t2 + lFinTag
   Loop
   
   FwFindXmlTag = -4 ' No hay tantas
   
End Function


' Permite encontrar la posición de inicio y término del contenido de un Tag en u texto XML
' Retorna la posición de inicio y en Rc() se puede obtener la posición de término.
' Tag: el elemento a buscar sin '<' ni '>'
' Indice: permite indicar que se busca el i-esimo tag. 0 o 1 es el primero.
' iPos: permite indicar desde donde buscar dentro del Xml. 0: indica no limite.
' fPos: permite indicar hasta donde buscar dentro del Xml. 0: indica no limite.
' ver 2010.06
Public Function FwGetXmlTag(ByVal xml As String, ByVal Tag As String, Optional ByVal Indice As Long = 0, Optional ByVal iPos As Long = 0, Optional fPos As Long = 0) As String
   Dim t0 As Integer, t1 As Long, t2 As Long, t3 As Long, i As Integer, lFld As Integer
   Dim Fld As String, Buf As String, l As Integer, Fin As Long, bFin As Boolean
   
   bFin = (fPos = -1)
   
   t1 = FwFindXmlTag(xml, Tag, Indice, iPos, fPos, t2)
   If t1 <= 0 Then
      gFwErr = t1
      Exit Function
   End If
   
   i = InStr(Tag, "|")
   If i > 0 Then
      Fld = Trim(Mid(Tag, i + 1))
      Tag = Trim(Left(Tag, i - 1))
   Else
      Fld = ""
      Tag = Trim(Tag)
   End If

   If t1 <= 0 Then
      gFwErr = -1
      Exit Function
   End If
   
   Buf = Mid(xml, t1, t2 - t1)
   
   If bFin Then
      fPos = t2
   End If
   
   If Fld = "" Then
      FwGetXmlTag = Buf
      Exit Function
   End If
   
   lFld = Len(Fld)

   l = 1
   t1 = InStr(1, Buf, Fld & "=", vbTextCompare)
   If t1 <= 0 Then
      l = 2
      t1 = InStr(1, Buf, Fld & " =", vbTextCompare)
      If t1 <= 0 Then
         gFwErr = -4
         Exit Function
      End If
   End If
   
   t1 = t1 + lFld + l
   If Mid(Buf, t1, 1) = """" Then
      t2 = InStr(t1 + 1, Buf, """", vbBinaryCompare)
      If t2 > 0 Then
         FwGetXmlTag = Mid(Buf, t1 + 1, t2 - (t1 + 1))
      Else
         FwGetXmlTag = ""
      End If
   Else
      t2 = InStr(t1, Buf, " ", vbBinaryCompare)
      If t2 > 0 Then
         FwGetXmlTag = Mid(Buf, t1, t2 - t1)
      Else
         FwGetXmlTag = Mid(Buf, t1)
      End If
   End If

End Function
' para por ejemplo <td><span>dato</span></td>
Function FwRemXmlTag(ByVal Fld As String, ByVal Tag As String) As String
   
   If InStr(Fld, "<" & Tag) > 0 Then
      FwRemXmlTag = FwGetXmlTag(Fld, Tag)
   Else
      FwRemXmlTag = Fld
   End If

End Function


Public Sub ResetDataChg(Frm As Form)
   On Error Resume Next
   Dim Ctrl As Control
   For Each Ctrl In Frm.Controls
'      Debug.Print Ctrl.Name
      Ctrl.DataChanged = False
   Next Ctrl

End Sub
Public Function ChkDataChg(Frm As Form) As Boolean
   On Error Resume Next
   Dim Ctrl As Control
   ChkDataChg = False
   For Each Ctrl In Frm.Controls
'      Debug.Print Ctrl.Name

      ERR = 0
      If Ctrl.DataChanged Then
         If ERR = 0 Then ' nos saltamos los que no tienen esta propiedad
            ChkDataChg = True
            Exit For
         End If
      End If
      
   Next Ctrl

End Function


Public Function FileSize(ByVal fname As String) As Long
   Dim Size As Long
   
   Size = -1

   On Error Resume Next
   Size = FileLen(fname)

   FileSize = Size
End Function


Public Function FileDate(ByVal fname As String) As Double
   Dim Dt As Double
   
   Dt = -1

   On Error Resume Next
   Dt = FileDateTime(fname)

   FileDate = Dt
End Function

' Ext: .jpg, .doc
Public Function GenTmpFile(ByVal Pref As String, ByVal Ext As String) As String
   Dim i As Long, fname As String

   For i = 1000 To 2000
      fname = Pref & "_" & Hex(i) & Ext

      If ExistFile(W.TmpDir & "\" & fname) = False Then
         GenTmpFile = W.TmpDir & "\" & fname
         Exit Function
      End If

   Next i

   fname = Pref & "_" & Format(Now, "dnnss") & Ext

   GenTmpFile = W.TmpDir & "\" & fname

End Function
' 24 oct 2014: pam: Se crea esta funcion porque el left con parámetro negativo se cae
Public Function FwLeft(ByVal Buf As String, ByVal lBuf As Long) As String

   If lBuf < 0 Then
      Exit Function
   Else
      FwLeft = Left(Buf, lBuf)
   End If

End Function

Public Function ChkFilled(Tx As TextBox, ByVal MinLen As Long) As Boolean

   Tx.Text = Trim(Tx.Text)
   If Len(Tx.Text) < MinLen And Tx.Enabled = True And Tx.Locked = False Then
      ChkFilled = False
   Else
      ChkFilled = True
   End If

End Function
Public Sub ColFilled(Tx As Control, Optional ByVal bFalta As Boolean = 0)

   If bFalta Then
      If gBkColOblig Then
         Tx.BackColor = gBkColOblig
      Else
         Tx.BackColor = &HC0FFFF
      End If
   Else
      Tx.BackColor = vbWindowBackground
   End If

End Sub
Public Function TxtFilled(Tx As TextBox, ByVal Msg As String, Optional lMin As Integer = 0) As Boolean

   Tx.Text = Trim(Tx.Text)
   If Len(Tx.Text) <= lMin Then
   
      If gBkColOblig Then
         Tx.BackColor = gBkColOblig
      End If
      
      If Len(Msg) > 0 Then
         MsgBox1 Msg, vbExclamation
      End If
      
      If Tx.Enabled Then
         Tx.SetFocus
      End If
      Exit Function
   End If

   Tx.BackColor = vbWindowBackground
   TxtFilled = True

End Function

Public Function CbSelected(Cb As Control, ByVal Msg As String) As Boolean

   If CbItemData(Cb) <= 0 Then
      Call ColFilled(Cb, True)
      MsgBox1 Msg, vbExclamation
      Cb.SetFocus
      Exit Function
   Else
      Call ColFilled(Cb)
      CbSelected = True
   End If

End Function

Public Function MkDirect(ByVal Path As String, Optional ByVal bOmitExist As Boolean = 1) As Long
   On Error Resume Next
   
   MkDir Path
   
   If bOmitExist = False Or ERR.Number <> ERR_PATHFILE Then ' 14 feb 2018: si ya existe no seria error
      MkDirect = ERR.Number
   End If
   
End Function

Public Function GetDebug() As Integer
   Dim Dbg As Integer
   
   Dbg = Val(GetCmdParam("Dbg", , "1"))
   If Dbg = 0 Then
      If FileSize(W.AppPath & "\Debug.txt") > 0 Then
         Dbg = 1
      End If
   End If

   GetDebug = Dbg
   
End Function

Public Sub MvControl(CtrFrm As Control, CtrTo As Control, Optional ByVal bSize As Boolean = True)

   CtrFrm.Left = CtrTo.Left
   CtrFrm.Top = CtrTo.Top
   
   If bSize Then
      CtrFrm.Width = CtrTo.Width
      CtrFrm.Height = CtrTo.Height
   End If

End Sub
Public Function DiasDelMes(ByVal AnoMes As Long) As Integer

   If AnoMes > 190001 Then ' yyyymm
      DiasDelMes = Day(DateSerial(AnoMes \ 100, (AnoMes Mod 100) + 1, 1) - 1)
   Else ' Juliano  73415 => 31 dic 20100
      DiasDelMes = Day(DateSerial(Year(AnoMes), month(AnoMes) + 1, 1) - 1)
   End If
   
End Function
' Suma o Resta meses a AnoMes
Public Function AnoMesAdd(ByVal AnoMes As Long, ByVal NMeses As Integer) As Long
   Dim Dt As Long
   
   Dt = DateSerial(AnoMes \ 100, (AnoMes Mod 100) + NMeses, 1)

   AnoMesAdd = Year(Dt) * 100& + month(Dt)

End Function

Public Function Dt2AnoMes(ByVal Dt As Long) As Long
   
   Dt2AnoMes = Year(Dt) * 100& + month(Dt)

End Function

Public Function LineCount(ByVal Fn As String, Optional ByVal MaxLin As Long = 0) As Long
   Dim Fd As Integer, Buf As String, l As Long

   On Error Resume Next
   Fd = FreeFile()
   Open Fn For Input As #Fd
   If ERR Then
      MsgErr Fn
      LineCount = -ERR.Number
      Exit Function
   End If
   
   l = 0
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      
      If l > MaxLin Then
         Close #Fd
         LineCount = -1
         Exit Function
      End If
      
   Loop
   Close #Fd

   LineCount = l

End Function

Public Function Str2Hex(ByVal Buf As String) As String
   Dim sTemp As String, c As Long

   sTemp = ""
   For c = 1 To Len(Buf)
       sTemp = sTemp & Right("0" & Hex(Asc(Mid(Buf, c, 1))), 2)
   Next
   Str2Hex = sTemp

End Function

Public Function Hex2Str(ByVal Buf As String) As String
   Dim sTemp As String, c As Long, v As Integer

   sTemp = ""
   For c = 1 To Len(Buf) Step 2
      v = Val("&H" & Mid(Buf, c, 2))
      sTemp = sTemp & Chr(v)
   Next
   Hex2Str = sTemp

End Function

Public Function Sign(ByVal Valor As Long) As Integer

   If Valor > 0 Then
      Sign = 1
   ElseIf Valor < 0 Then
      Sign = -1
   Else
      Sign = 0
   End If

End Function

Public Sub Blink(lb As Label, Tm As Timer)

   Tm.Enabled = False
   If lb.visible Then
      Tm.Interval = 300
      lb.visible = False
   Else
      Tm.Interval = 700
      lb.visible = True
   End If
   Tm.Enabled = True

End Sub

Public Function PathFromFilename(ByVal Fn As String) As String
   Dim i As Integer
   
   i = rInStr(Fn, "\")
   If i > 0 Then
      PathFromFilename = Left(Fn, i - 1)
   Else
      PathFromFilename = ""
   End If

End Function

' Para glosas que deben ser sólo una linea
Public Function RemoveSpcChars(ByVal Buf As String) As String

   RemoveSpcChars = ReplaceStr(ReplaceStr(Buf, vbCr, ""), vbLf, "")

End Function
' 30 ago 2019: Para no tener problema con excel
Public Function RemoveNoPrtChars(ByVal Buf As String, Optional ByVal bNoQuotes As Boolean = 0) As String
   Dim i As Integer, ch As String

   For i = 0 To 31
      ch = Chr(i)
      Buf = Replace(Buf, ch, " ")
   Next i

   If bNoQuotes Then
      Buf = Replace(Buf, """", " ")  'FCA 9 abr 2020, para no tener problemas con el Excel
   End If
   
   For i = 127 To 144
      ch = Chr(i)
      Buf = Replace(Buf, ch, " ")
   Next i

   For i = 153 To 160
      ch = Chr(i)
      Buf = Replace(Buf, ch, " ")
   Next i

   For i = 164 To 165
      ch = Chr(i)
      Buf = Replace(Buf, ch, " ")
   Next i

   For i = 167 To 190
      ch = Chr(i)
      Buf = Replace(Buf, ch, " ")
   Next i

   RemoveNoPrtChars = Buf

End Function

Public Sub GenError(ByVal bGenErr As Boolean, Optional ByVal bRnd As Boolean = 1)
   Dim i As Integer
   
   If bGenErr Then
      ERR.Raise 6, "Overflow"
      On Error GoTo 0
      i = 32500 + IIf(bRnd, 10000, Rnd() * 10000)
   End If

End Sub
' Enmascara un campo para que sea CSV, si trae un ; lo pone entre " y si viene con una " la cambia por "" y lo pone entre "
Public Function FldCsv(ByVal Fld As String, Optional ByVal Sep As String = ";") As String

   If InStr(Fld, """") > 0 Or InStr(Fld, Sep) > 0 Then
      FldCsv = """" & ReplaceStr(Fld, """", """""") & """"
   Else
      FldCsv = Fld
   End If

End Function

' Convierte Tab Separated por CSV (separado por ;)
Public Function Tab2xCSV(ByVal Buf As String) As String
   Dim p As Long, Fld As String, lBuf As Long, Out As String
   
   p = 1
   lBuf = Len(Buf)
   Out = ""
   Do While p <= lBuf

      If p > 1 Then
         Out = Out & ";"
      End If
      
      Fld = NextField2(Buf, p)

'      If InStr(Fld, """") > 0 Or InStr(Fld, ";") > 0 Then
'         Fld = """" & ReplaceStr(Fld, """", """""") & """"
'      End If

      Out = Out & FldCsv(Fld, ";")

   Loop

   Tab2xCSV = Out

End Function
' 7 nov 2017: concatena el archivo FnRead a continuación del archivo FnWrite
Public Function FileCat(ByVal FnWrite As String, ByVal FnRead As String) As Long
   Dim FdW As Integer, FdR As Integer, Buf As String
   
   
   On Error Resume Next
   
   FdW = FreeFile()
   Open FnWrite For Append As #FdW
   If ERR Then
      FileCat = ERR.Number
      Exit Function
   End If
   
   FdR = FreeFile()
   Open FnRead For Input As #FdR
   If ERR Then
      FileCat = ERR.Number
      Exit Function
   End If
   
   Do Until EOF(FdR)
      Line Input #FdR, Buf
   
      Print #FdW, Buf
   Loop
   
   Close #FdW
   Close #FdR

End Function

Public Sub RemoveNL(Buf As String, Optional ByVal Rep As String = "")

   Buf = Replace(Buf, vbCrLf, Rep)
   Buf = Replace(Buf, vbCr, Rep)
   Buf = Replace(Buf, vbLf, Rep)

End Sub
Public Function TieneAcentos(ByVal Buf As String) As Boolean
   Dim Chrs As String, bTiene As Boolean, i As Integer
   
   Chrs = "áéíóúñü"
   
   For i = 1 To Len(Chrs)
      If InStr(1, Buf, Mid(Chrs, i, 1), vbTextCompare) > 0 Then
         TieneAcentos = True
         Exit Function
      End If
   Next i

End Function
Public Function EsAcento(ByVal ch As String) As Boolean
   Dim Chrs As String
   
   Chrs = "áéíóúñü"

   If InStr(1, Chrs, ch, vbTextCompare) > 0 Then
      EsAcento = True
   End If
   
End Function

' Para quitar acentos
Public Function ToAscii(ByVal Buf As String) As String
   Dim i As Integer, c As String * 1, j As Integer
   Dim s As String, o As String
   
   s = "áéíóúñüÁÉÍÓÚÑÜ"
   o = "aeiounuAEIOUNU"
   
   For i = 1 To Len(Buf)
      c = Mid(Buf, i, 1)
      
      If Asc(c) < 32 Or Asc(c) > 127 Then
         j = InStr(s, c)
         If i > 0 Then
            c = Mid(o, j)
         Else
            c = "-"
         End If
         
         Mid(Buf, i, 1) = c
      
      End If
   Next i
      
   ToAscii = Buf

End Function

Public Function IsAlpha(ByVal Buf As String) As Boolean
   Dim iLen As Long
   Dim i As Long
   Dim sChar As String * 1
   
   Buf = Trim(Buf)
   iLen = Len(Buf)
   If iLen > 0 Then
      For i = 1 To iLen
         sChar = Mid(Buf, i, 1)
         If Not sChar Like "[A-Z a-z]" Then
            Exit Function
         End If
      Next
      
      IsAlpha = True
   End If
    
End Function

Public Function IsAlphaNum(ByVal Buf As String) As Boolean
   Dim iLen As Long
   Dim i As Long
   Dim sChar As String * 1
   
   Buf = Trim(Buf)
   iLen = Len(Buf)
   If iLen > 0 Then
      For i = 1 To iLen
         sChar = Mid(Buf, i, 1)
         If Not sChar Like "[0-9A-Z a-z]" Then
            Exit Function
         End If
      Next
      
      IsAlphaNum = True
   End If
        
End Function
' 23 nov 2018: no funciona
Public Function LockListbox(ls As ListBox, ByVal bLock As Boolean) As Boolean
   Dim Style As Long, Rc As Long
   
   Style = GetWindowLong(ls.hWnd, GWL_STYLE)

   If bLock Then
      Style = Style Or LBS_NOSEL
   Else
      Style = Style Xor LBS_NOSEL
   End If

   Rc = SetWindowLong(ls.hWnd, GWL_STYLE, Style)

   Rc = GetWindowLong(ls.hWnd, GWL_STYLE)

   LockListbox = (Rc = Style)

End Function


Public Function GetLastDay(ByVal AnoMes As Long) As Integer

   GetLastDay = Day(DateSerial(AnoMes \ 100, (AnoMes Mod 100) + 1, 1) - 1)

End Function

' Para que el On Error no afecte el resto
Public Function OpenTxFile(ByVal Filename As String, Optional ByVal bInput As Boolean = 1, Optional ByVal bMsg As Boolean = 1) As Integer
   Dim Fd As Integer

   On Error Resume Next
   
   Fd = FreeFile()
   
   If bInput Then
      Open Filename For Input As #Fd
   Else
      Open Filename For Output As #Fd
   End If
   
   If ERR.Number Then
      If bMsg Then
         MsgBox1 "Error al abrir el archivo" & vbCrLf & Filename & vbCrLf & "Err " & ERR.Number & ", " & ERR.Description, vbExclamation
      End If
      
      OpenTxFile = -1
   Else
      OpenTxFile = Fd
   End If
   
End Function

' Para archivos no tan grandes
Public Function ReadTxFile(ByVal Filename As String, Optional ByVal bMsg As Boolean = 1) As String
   Dim Fd As Integer, Buf As String, Texto As String
   
   Fd = OpenTxFile(Filename, True, bMsg)
   If Fd <= 0 Then
      Exit Function
   End If

   ReadTxFile = Input(LOF(Fd), Fd)

   Close #Fd
   
End Function
' PBar = ProgressBar
Public Sub SetProgressBar(PBar As Control, ByVal Value As Integer)
   
   If PBar Is Nothing Then
      Exit Sub
   End If
   
   If StrComp(TypeName(PBar), "ProgressBar", vbTextCompare) Then
      Debug.Print "SetProgressBar: el control no es una ProgressBar."
      Exit Sub
   End If
   
   If Value > PBar.Max Then
      PBar.Max = Value
   End If
   
   On Error Resume Next
   PBar.Value = Value
   DoEvents

End Sub
' https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa264975(v=vs.60)?redirectedfrom=MSDN
Public Function vbErrMsg(ByVal Erc As Integer, Optional ByVal Msg As String) As String

   Select Case Erc
      Case 6:
         vbErrMsg = "Desbordamiento"
      Case 6:
         vbErrMsg = "Indice fuera de rango"
      Case 11:
         vbErrMsg = "División por cero"
      Case 13:
         vbErrMsg = "Los tipos no calzan"
      Case 51:
         vbErrMsg = "Error interno"
      Case 52:
         vbErrMsg = "Descriptor o nombre de archivo inválido"
      Case 53:
         vbErrMsg = "Archivo no encontrado"
      Case 53:
         vbErrMsg = "Archivo ya abierto"
      Case 54:
         vbErrMsg = "Modo de archivo inválido"
      Case 58:
         vbErrMsg = "Archivo ya existe"
      Case 61:
         vbErrMsg = "Disco lleno"
      Case 62:
         vbErrMsg = "Lectura más allá del final del archivo"
      Case 67:
         vbErrMsg = "Demasiados archivos"
      Case 68:
         vbErrMsg = "Unidad o dispositivo no disponible"
      Case 70:
         vbErrMsg = "Permiso denegado"
      Case 71:
         vbErrMsg = "El disco no está listo"
      Case 75:
         vbErrMsg = "Error de acceso a ruta o archivo"
      Case 76:
         vbErrMsg = "Ruta no encontrada"
      Case 91:
         vbErrMsg = "Objeto o variable no asignada"
      Case 321:
         vbErrMsg = "Formato de archivo inválido"
      Case Else:
         vbErrMsg = Msg
   End Select

End Function

' Para los Subject con acentos
Public Function EncodeSubject(ByVal Buf As String) As String
   Dim i As Integer, ch As String, Enc As String, bNonAscii As Boolean, c As Integer
   
   For i = 1 To Len(Buf)
      ch = Mid(Buf, i, 1)
      c = Asc(ch)
      If c < 32 Or c > 126 Then
         Enc = Enc & "=" & Right("00" & Hex(c), 2)
         bNonAscii = True
      Else
         Enc = Enc & ch
      End If
   Next i

   If bNonAscii Then
      EncodeSubject = "=?ISO-8859-1?Q?" & Enc & "?="
   Else
      EncodeSubject = Buf
   End If

End Function
