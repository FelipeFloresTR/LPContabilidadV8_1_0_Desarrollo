Attribute VB_Name = "FAIRDLL32"
Option Explicit

Public gInFairware As Boolean

Type FwNivel_t
   id        As Long
   Desc      As String
   GenCode   As Byte
End Type

Type AppCode_t
   Rc          As Integer
   
   Name        As String   ' nombre interno de la aplicación para verificar si está activo o no
   Code        As Long     ' Código de la aplicación
   TVerif      As Byte
   Rut         As String
   PcCode      As String
   Title       As String   ' Título para el Caption
   IniFile     As String
   CfgFile     As String
   Demo        As Boolean
   bTilt       As Boolean  ' Para marcar si es un equipo eliminado que está inscrito
   NivProd     As Long
   UnReg       As Boolean
   TxInsc1     As String   ' texto de la ventana de Inscripción
   TxInsc2     As String   ' texto de la ventana de Inscripción
   emailSop    As String   ' email de soporte
   emailInfo   As String   ' email de contacto  soporte@fairware.cl
   TelInfo     As String   ' email de contacto  soporte@fairware.cl
   Contacto    As String   ' Fairware
   Nivel(20)   As FwNivel_t  ' desde 0 hasta Desc = ""
                              ' Si el primero viene con Desc "" se asume id=1 y único
                              
   NivDef      As Long  ' nivel por defecto, el más limitado, cuando no tiene
   FUsoVersion As Long     ' Fecha primer uso de esta version
   
   Msg         As String
   MinMsg      As Integer  ' Cada cuantos minutos muestra el Msg
   
   GenCod      As Integer  ' Si GenCod=2 y Nivel > 5, entonces usar GenCod2
   
   txDemo      As String
End Type
Public gAppCode As AppCode_t

Type CCMB_T
    B1                  As String * 1
    B2                  As String * 1
    B3                  As String * 1
    B4                  As String * 1
    Func                As String * 1
    Rcodelo             As String * 1
    Rcodehi             As String * 1
    DriveNo             As String * 1
    Dir                 As Long
    Vers                As String * 2
    ProductSerialNumber As Integer
    ProductCode         As String * 9
    ProgramName         As String * 13
    CCSerialNumber      As Integer
    Master              As String * 1
    DriveType           As String * 1
    Copies              As Integer
    InitCopies          As Integer
    Uses                As Integer
    InitUses            As Integer
    ExpiryDay           As String * 1
    ExpiryMonth         As String * 1
    ExpiryYear          As Integer
    Feature             As Integer
    MaxNetUsers         As Integer
    SecureMsg           As String * 257
    UpdateNumberlo      As String * 1
    UpdateNumberhi      As String * 1
    Flags1lo            As String * 1
    Flags1hi            As String * 1
    NetUserData         As String * 4
    DosTime             As String * 4
    Maxdayslo           As String * 1
    Maxdayshi           As String * 1
    Res2                As String * 4
    ExtendedError       As String * 6
    ResExpand           As String * 173

    'B1          As String * 1
    'B2          As String * 1
    'B3          As String * 1
    'B4          As String * 1
    'Func        As String * 1
    'Rcode       As Integer
    'DriveNo     As String * 1
    'DirOffset   As Integer
    'DirSeg      As Integer
    'OtherBits1  As String * 26
    'CCSerialNumber As Integer
    'OtherBits2  As String * 277
    'Flags1      As Integer
    'Remainder   As String * 193
End Type

Type Tm_t
    Sec   As Integer  ' Seconds after the minute - [0,59]
    Min   As Integer  ' Minutes after the hour - [0,59]
    Hour  As Integer  ' Hours since midnight - [0,23]
    MDay  As Integer  ' Day of the month - [1,31]
    Mon   As Integer  ' Months since January - [0,11]
    Year  As Integer  ' Years since 1900 : Year(Now) - 1900
    WDay  As Integer  ' Days since Sunday - [0,6]   : WeekDay(Now) - 1
    YDay  As Integer  ' Days since January 1 - [0,365] : Int( Now - Dateserial(Year(Now),1,1)) + 1
    IsDST As Integer  ' Daylight-saving-time flag
End Type


' Funciones de registro dependientes del N° de serie del equipo
Declare Sub FwPcCode Lib "FairDll32.dll" Alias "Fw00001" (ByVal Seed As Long, ByVal PcCode As String)
Declare Function FwCheckKey Lib "FairDll32.dll" Alias "Fw00003" (ByVal Seed As Long) As Long
Declare Function FwSetUserCode Lib "FairDll32.dll" Alias "Fw00004" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long
Declare Function FwDelUserCode Lib "FairDll32.dll" Alias "Fw00005" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long

' Funciones de registro NO dependientes del N° de serie del equipo
'Declare Sub FwPcCode2 Lib "FairDll32.dll" Alias "Fw00011" (ByVal Seed As Long, ByVal PcCode As String)
'Declare Function FwCheckKey2 Lib "FairDll32.dll" Alias "Fw00012" (ByVal Seed As Long) As Long
'Declare Function FwSetUserCode2 Lib "FairDll32.dll" Alias "Fw00013" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long
'Declare Function FwDelUserCode2 Lib "FairDll32.dll" Alias "Fw00014" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long

' Funciones de registro NO dependientes del N° de serie del equipo, W. Vista
Declare Sub FwPcCode3 Lib "FairDll32.dll" Alias "Fw00011" (ByVal Seed As Long, ByVal PcCode As String)
Declare Function FwCheckKey3 Lib "FairDll32.dll" Alias "Fw00012" (ByVal Seed As Long) As Long
Declare Function FwSetUserCode3 Lib "FairDll32.dll" Alias "Fw00013" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long
Declare Function FwDelUserCode3 Lib "FairDll32.dll" Alias "Fw00014" (ByVal Seed As Long, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long

' Funciones
Declare Sub FwInit Lib "FairDll32.dll" Alias "Fw00018" (ByVal Buf As String, ByVal Dat As Long)
Declare Function FwUserCode Lib "FairDll32.dll" Alias "Fw00002" (ByVal Seed As Long, ByVal PcCode As String, ByVal Rut As String, ByVal Level As Long, ByVal UserCode As String) As Long
Declare Function FwVersion Lib "FairDll32.dll" Alias "Fw00006" (ByVal sVersion As String, ByVal Ln As Long) As Long
Declare Function FwCCDLL Lib "FairDll32.dll" Alias "Fw00008" (ByVal Path As String, cc As CCMB_T) As Long

' Asigna la fecha al archivo
Declare Function FwUTime Lib "FairDll32.dll" Alias "Fw00009" (ByVal PName As String, Tm As Tm_t) As Long

' Obtiene la fecha del archivo
Declare Function FwFdTime Lib "FairDll32.dll" Alias "Fw00010" (ByVal Fd As Integer, Tm As Tm_t) As Long

Declare Sub FwMemSet Lib "FairDll32.dll" Alias "Fw00015" (Buf As Any, ByVal ch As Integer, ByVal BufLen As Integer)
Declare Sub FwMemCpy Lib "FairDll32.dll" Alias "Fw00016" (Dest As Any, Src As Any, ByVal BufLen As Integer)

#If FWREG = 0 Then
Public Sub FwRegist()
   Dim Frm As FrmInscripcion
   Dim FDemo As FrmDemo

   Call FwInitAppCode
   
   If gAppCode.Demo Then
      ' "Producto NO Inscrito"

      Set FDemo = New FrmDemo
      FDemo.Show vbModal
      Set FDemo = Nothing
      
      If gRc.Rc = vbCancel Then
         End
      ElseIf gRc.Rc = vbOK Then
   
         Set Frm = New FrmInscripcion
         Frm.Show vbModal
         Set Frm = Nothing
         
         If gAppCode.Rc <> vbOK Then
            End
         End If
         
      End If
   Else
      gAppCode.Demo = False

   End If

End Sub
#End If

#If FWREG = 0 Then
Public Sub FwUnRegist()
   Dim Frm As FrmInscripcion
   
   gAppCode.UnReg = True

   Set Frm = New FrmInscripcion
   Frm.Show vbModal
   Set Frm = Nothing

End Sub
#End If

Public Function FwDver(ByVal Code As String) As String
   Dim x As Long, i As Integer, m As Integer, j As Integer
   
   If Len(Code) < 10 Then
      FwDver = ""
      Exit Function
   End If
   
   x = 0
   For i = 1 To 10
      j = (i Mod 3) + 1
      x = x + Asc(Mid(Code, i, 1)) * i * j * 3
   Next i

   m = x Mod 23

   FwDver = Chr(Asc("C") + m)

End Function
' Codigo de PC + verificador
Public Function FwGetPcCode() As String
   Dim Buf As String, Rc As Long

   If gAppCode.Code = 0 Then
      Debug.Print "*** falta gAppCode.Code ***"
   End If
   
   Buf = Space(30)
   Call FwPcCode3(gAppCode.Code, Buf)
   Buf = Trim(Left(Buf, StrLen(Buf)))
   If Buf <> "" Then
      FwGetPcCode = Left(Buf, 10) & FwDver(Left(Buf, 10))
   Else
      Debug.Print "DllError: " & Err.LastDllError & ", " & GetLastSystemError(Err.LastDllError)
   End If

End Function
' Codigo de PC + verificador
Public Function FwGetVer() As String
   Dim Buff As String * 100
   Dim Ver As Long

   Ver = FwVersion(Buff, 99)
   'FwGetVer = Left(Buff, StrLen(Buff))

   FwGetVer = Ver \ (256# * 256#) & "." & Ver Mod (256# * 256#)

End Function

Public Function FwUnRegister(Optional ByVal UserCode As String = "") As Long
   Dim Buf30 As String * 20
   Dim Rc As Long

   FwUnRegister = -1
   If FwDelUserCode3(gAppCode.Code, "", gAppCode.NivProd, UserCode) = 0 Then
      FwUnRegister = 0
      Exit Function
   End If
   
   ' *** Este código no debería ser necesario, pero...
   
   If FwVersion("", 0) >= &H20004 Then ' *** por ahora
      Call FwInit("", 9126670) ' para que permita obtener el código para eliminar
   End If

   Rc = FwUserCode(gAppCode.Code, Left(FwGetPcCode(), 10), "", gAppCode.NivProd, Buf30)
   
   If Rc = 0 Then
      Rc = FwDelUserCode3(gAppCode.Code, "", gAppCode.NivProd, Left(Buf30, 10))
   End If
   
   FwUnRegister = Rc
   
End Function

Public Sub FwInitAppCode()
   
   Debug.Print "Demo=" & FwEncrypt1(Space(13) & "DEMO" & Space(11), 76172)
   gAppCode.txDemo = Trim(FwDecrypt1("23672C7239814A945F2B78C66689B563368A5F358C643D97724E2B89", 76172))

   gAppCode.Demo = True
   gAppCode.bTilt = False
   
   If gAppCode.emailSop = "" Then
      gAppCode.emailSop = "soporte@fairware.cl"
   End If
   
   If gAppCode.emailInfo = "" Then
      gAppCode.emailInfo = "info@fairware.cl"
   End If

   If gAppCode.TelInfo = "" Then
      gAppCode.TelInfo = "(56 2) 2212 1594"
   End If

   If gAppCode.Contacto = "" Then
      gAppCode.Contacto = "Fairware"
   End If

   If gAppCode.CfgFile = "" Then
      Debug.Print "**** OJO: Falta asignar el archivo CFG."
      Beep
      Debug.Assert ""
      End
   End If

   If gAppCode.IniFile = "" Then
      Debug.Print "**** OJO: Falta asignar el archivo INI."
      Beep
      Debug.Assert ""
      End
   End If

   If gAppCode.Code = 0 Then
      Debug.Print "**** OJO: Falta asignar el código (SEED) de la Aplicación."
      Beep
      Debug.Assert ""
      End
   End If

   gAppCode.PcCode = FwGetPcCode()
   If gAppCode.PcCode = "" Then
      Debug.Print "**** OJO: Falta la FairDll32.dll junto al ejecutable o llamar a FwInit."
      Beep
   End If


   Call FwGetLicRut
   
'   gAppCode.Rut = GetIniString(gAppCode.CfgFile, "Config", "RUT")
   If gAppCode.Rut = "" Then
      Debug.Print "**** OJO: Falta asignar el código RUT del cliente."
      Beep
   End If


   ' Primero vemos si estaba inscrito a la antigua
   gAppCode.NivProd = FwCheckKey(gAppCode.Code)
   If gAppCode.NivProd <= 0 Then
      gAppCode.NivProd = FwCheckKey3(gAppCode.Code)
   End If
   
   If gAppCode.NivProd <= 0 Then
      gAppCode.NivProd = 1
      gAppCode.Demo = True

   Else
      gAppCode.Demo = False
      Call SetIniString(gAppCode.IniFile, "Config", "PCName", GetComputerName())
   
   End If

End Sub

Public Function FwGetErr(ByVal ErrCode As Long) As String

   Select Case ErrCode
      Case -121:
         FwGetErr = "No Init detected"
      Case -101:
         FwGetErr = "No match"
      Case -121:
         FwGetErr = "Bad param"
      Case Else:
         FwGetErr = "Unknown error"
         
   End Select

End Function

Private Function FwKeyRut(ByVal Rut As String) As Long
   ' La key no puede depender del PC porque es para todos
   FwKeyRut = GenClave2("[" & gAppCode.Name & "#" & Rut & "]", 74213)       ' 10 ago 2012
End Function

Public Sub FwSetLicRut()
   Dim Key As Long
   
   If gAppCode.Rut <> "" Then
      Key = FwKeyRut(gAppCode.Rut)      ' 10 ago 2012
      
      Call SetIniString(gAppCode.CfgFile, "Config", "LicRUT", gAppCode.Rut)
      Call SetIniString(gAppCode.CfgFile, "Config", "ChkRUT2", Key)
      Call SetIniString(gAppCode.CfgFile, "Config", "ChkRUT", vbNullString)
   End If
End Sub

Public Sub FwGetLicRut()
   Dim Rut As String, Key1 As Long, Key2 As Long, bSet As Boolean

   Rut = GetIniString(gAppCode.CfgFile, "Config", "LicRUT") ' rut de la licencia
   Key2 = Val(GetIniString(gAppCode.CfgFile, "Config", "ChkRUT2"))
   
   If Rut = "" Or Key2 = 0 Then ' si no tenía asumimos el RUT de la oficina
      Rut = GetIniString(gAppCode.CfgFile, "Config", "RUT")
      bSet = True ' no tiene ChkRUT2
   End If
   
   Key1 = FwKeyRut(Rut)   ' 10 ago 2012
   
   If Key1 = Key2 Or Key2 = 0 Then
      gAppCode.Rut = Rut
      
      If bSet Then
         Call FwSetLicRut
      End If
      
   Else
      gAppCode.Rut = "X" & Rut ' mayusculas
      
      If bSet And Now < DateSerial(2015, 3, 16) Then  ' solo por un plazo fijo
         gAppCode.Rut = Rut
         Call FwSetLicRut
      End If
   
   End If

End Sub

