VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEquiposAut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración Licencias de la Red"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   Icon            =   "FrmEquiposAut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_nLic 
      Height          =   315
      Left            =   9360
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "0"
      Top             =   900
      Width           =   495
   End
   Begin VB.TextBox Tx_Oficina 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   480
      Width           =   6495
   End
   Begin VB.TextBox Tx_NetCode 
      Height          =   315
      Left            =   3480
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1320
      Width           =   2115
   End
   Begin VB.ComboBox Cb_Nivel 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   3015
   End
   Begin VB.TextBox Tx_Cod 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   10140
      TabIndex        =   17
      Text            =   "no visible"
      Top             =   3060
      Width           =   1875
   End
   Begin VB.CommandButton Bt_DelPC 
      Caption         =   "Eliminar equipo..."
      Height          =   315
      Left            =   10140
      TabIndex        =   11
      Top             =   4800
      Width           =   1995
   End
   Begin VB.TextBox Tx_RUT 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "La licencia queda asociada a este RUT."
      Top             =   480
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   10140
      TabIndex        =   8
      Top             =   1260
      Width           =   1995
   End
   Begin VB.CommandButton Bt_Lic 
      Caption         =   "Leer Licencia de Uso y Garantía Limitada"
      Height          =   675
      Left            =   10140
      TabIndex        =   13
      Top             =   5940
      Width           =   1995
   End
   Begin VB.CheckBox Ck_Lic 
      Caption         =   "Acepto las Condiciones de la Licencia de Uso"
      Height          =   615
      Left            =   10140
      TabIndex        =   12
      Top             =   5220
      Width           =   1935
   End
   Begin VB.CommandButton Bt_Send 
      Caption         =   "Solicitar código de red..."
      Height          =   615
      Left            =   10140
      TabIndex        =   9
      ToolTipText     =   "Prepara un correo con la información necesaria."
      Top             =   1680
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Bt_Copy 
      Caption         =   "Licenciar Online"
      Height          =   315
      Left            =   10140
      TabIndex        =   10
      Top             =   2400
      Width           =   1995
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10140
      TabIndex        =   7
      Top             =   840
      Width           =   1995
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   10140
      TabIndex        =   6
      Top             =   480
      Width           =   1995
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   1440
      TabIndex        =   5
      Top             =   1860
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   40
      Cols            =   4
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   360
      Picture         =   "FrmEquiposAut.frx":000C
      Top             =   540
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<=  Licencias:"
      Height          =   195
      Index           =   2
      Left            =   8280
      TabIndex        =   24
      Top             =   960
      Width           =   990
   End
   Begin VB.Label La_nPCs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Left            =   7740
      TabIndex        =   23
      Top             =   900
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Equipos:"
      Height          =   195
      Index           =   1
      Left            =   7080
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.Label La_Ver 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10140
      TabIndex        =   21
      Top             =   6660
      Width           =   1995
   End
   Begin VB.Label La_InfoLic 
      Caption         =   "Marque en la columna Licencia todos los equipos que se incluirán en la solicitud."
      Height          =   1035
      Left            =   10140
      TabIndex        =   20
      Top             =   3540
      Width           =   1995
   End
   Begin VB.Label La_Niv 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
      Height          =   315
      Left            =   4980
      TabIndex        =   18
      Top             =   900
      Width           =   615
   End
   Begin VB.Label La_NetCode 
      AutoSize        =   -1  'True
      Caption         =   "&Código de Licencia de Red:"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   1380
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Nivel:"
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RUT:"
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   16
      Top             =   540
      Width           =   390
   End
   Begin VB.Label La_Aut 
      Alignment       =   2  'Center
      Caption         =   "Autorizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5820
      TabIndex        =   15
      Top             =   1380
      Width           =   2895
   End
End
Attribute VB_Name = "FrmEquiposAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxLic = 200

Private Const K_SEED = 7194171
Private Const K_VX = 7876345

Private Const C_PC = 0
Private Const C_MAC = 1
Private Const C_CODPC = 2
Private Const C_ULT = 3
Private Const C_CLIC = 4   ' Con licencia
Private Const C_LIC = 5   ' Licencia
Private Const C_IDX = 6

Private Const LASTCOL = C_IDX

Private lModif As Boolean
Private lAdm As Boolean
Private lnPCs As Integer

Private Sub SetupForm()

   Grid.Cols = LASTCOL + 1

   Call FGrSetup(Grid)
   
   Grid.TextMatrix(0, C_PC) = "Nombre PC"
   Grid.TextMatrix(0, C_MAC) = "MAC"
   Grid.TextMatrix(0, C_CODPC) = "Código PC"
   Grid.TextMatrix(0, C_ULT) = "Último uso"
   Grid.TextMatrix(0, C_CLIC) = "Con licencia"
   Grid.TextMatrix(0, C_LIC) = "Licencia"

   Grid.ColWidth(C_PC) = 1800
   Grid.ColWidth(C_MAC) = 1600
   Grid.ColWidth(C_CODPC) = 1500
   Grid.ColWidth(C_ULT) = 1100
   Grid.ColWidth(C_CLIC) = 1000
   Grid.ColWidth(C_LIC) = 1000
   Grid.ColWidth(C_IDX) = 0

   Grid.ColAlignment(C_PC) = flexAlignLeftCenter
   Grid.ColAlignment(C_MAC) = flexAlignCenterCenter
   Grid.ColAlignment(C_CODPC) = flexAlignCenterCenter
   Grid.ColAlignment(C_ULT) = flexAlignRightCenter
   Grid.ColAlignment(C_CLIC) = flexAlignCenterCenter
   Grid.ColAlignment(C_LIC) = flexAlignCenterCenter

End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Copy_Click()
   Dim Body As String
   
   If Trim(Tx_RUT) = "" Or Val(Tx_RUT) = 0 Then
      MsgBox1 "Debe ingresar el RUT de su empresa.", vbExclamation
      Exit Sub
   End If

   If ValidRut(Tx_RUT) = False Then
      MsgBox1 "El RUT ingresado es inválido.", vbExclamation
      Call Tx_RUT.SetFocus
      Exit Sub
   End If

   If Ck_Lic.Value = 0 Then
      MsgBox1 "Antes de solicitar su Código de Red debe aceptar las Condiciones de la Licencia de Uso.", vbExclamation
      Call Ck_Lic.SetFocus
      Exit Sub
   End If

   If Cb_Nivel.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un nivel.", vbExclamation
      Call Cb_Nivel.SetFocus
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   DoEvents

   Body = NewGenBody()

   If Body <> "" Then
   
      Clipboard.Clear
      Clipboard.SetText Body
      
      MsgBox1 "Hubo un problema para licenciar Online... " & vbCrLf & "Los datos fueron copiados al portapapeles." & vbCrLf & "Pegue estos datos en un correo dirigido a " & Trim(gAppCode.emailSop), vbInformation
   End If
   
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_DelPC_Click()
   Dim Row As Integer, PC As String, Mac As String, Cod As String
   Dim Q1 As String, Rc As Long, i As Integer
   
   Row = Grid.Row
   PC = Trim(Grid.TextMatrix(Row, C_PC))
   
   If PC = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If UCase(Trim(Grid.TextMatrix(Row, C_CLIC))) = "SI" Then
      MsgBox1 "No puede eliminar un equipo que está autorizado.", vbExclamation
      Exit Sub
   End If
   
   Mac = Trim(Grid.TextMatrix(Row, C_MAC))
   Cod = Trim(Grid.TextMatrix(Row, C_CODPC))

   If MsgBox1("Se eliminará el equipo:" & vbLf & vbLf & "PC: " & PC & vbLf & "MAC: " & Mac & vbLf & "Cód. PC: " & Cod & vbLf & vbLf & "¿ Desea continuar ?", vbYesNo Or vbExclamation Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   i = Val(Grid.TextMatrix(Row, C_IDX))
   Call SetIniString(gLicFile, PC_EQUIP, PC_NOM & i, "")
   Call SetIniString(gLicFile, PC_EQUIP, PC_COD & i, "")
   Call SetIniString(gLicFile, PC_EQUIP, PC_MAC & i, "")
   Call SetIniString(gLicFile, PC_EQUIP, PC_AUT & i, "")
   
   Grid.RemoveItem Row
   
End Sub

Private Sub Bt_Lic_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Licencia1.wri"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.rtf"
      Rc = ExistFile(Buf)
   End If
   
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.pdf"
      Rc = ExistFile(Buf)
   End If
      
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.htm"
      Rc = ExistFile(Buf)
   End If

   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene la licencia de uso de la aplicación." & vbCrLf & "Por favor contáctese con su proveedor para conseguirlo antes de solicitar el código.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene la licencia de uso y garantía de la aplicación." & vbCrLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub bt_OK_Click()
   Dim Code As String, Info As String
   
   If Trim(Tx_RUT) = "" Or Val(Tx_RUT) = 0 Then
      MsgBox1 "Ingrese el RUT de su empresa.", vbExclamation
      Exit Sub
   End If
   
   If ValidRut(Tx_RUT) = False Then
      MsgBox1 "El RUT ingresado es inválido.", vbExclamation
      Exit Sub
   End If

   If Cb_Nivel.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un nivel.", vbExclamation
      Call Cb_Nivel.SetFocus
      Exit Sub
   End If

'   If Val(La_nPCs) > Val(Tx_nLic) Then
'      MsgBox1 "La cantidad de equipos debe ser menor o igual que la cantidad de licencias.", vbExclamation
'      Call Cb_Nivel.SetFocus
'      Exit Sub
'   End If

   Info = GenInfo()
   If Info = "" Then
      Exit Sub
   End If
      
   If W.InDesign Then
      Tx_Cod = GenCodigo(UCase(Info))
   End If
      
   Code = Trim(UCase(Tx_NetCode))
   If Code = "" Then
      MsgBox1 "Debe ingresar el código de red.", vbExclamation
      Exit Sub
   End If
   
   If IsValidCode(Code) = False Then
      MsgBox1 "El código ingresado no es válido, revise las letras.", vbExclamation
      Exit Sub
   End If
   
   If Code <> GenCodigo(UCase(Info)) Then
      MsgBox1 "El código ingresado no corresponde a los equipos, al nivel seleccionados o a la cantidad de licencias.", vbExclamation
      Exit Sub
   End If

   Call SaveAll
   
   ' Nueva inscripción
   If CheckInscPC() Then
      MsgBox1 "Felicitaciones, Licenciamiento exitoso, gracias por preferir Thomson Reuters.", vbInformation
      MsgBox1 "El sistema se cerrara Favor volver a ingresar.", vbInformation
      End
   End If
   
   Unload Me
   
End Sub

Private Sub Bt_Send_Click()
   Dim Buf As String, Rc As Long, Info As String

   If Trim(Tx_RUT) = "" Or Val(Tx_RUT) = 0 Then
      MsgBox1 "Debe ingresar el RUT de su empresa.", vbExclamation
      Exit Sub
   End If

   If ValidRut(Tx_RUT) = False Then
      MsgBox1 "El RUT ingresado es inválido.", vbExclamation
      Call Tx_RUT.SetFocus
      Exit Sub
   End If

   If Cb_Nivel.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un nivel.", vbExclamation
      Call Cb_Nivel.SetFocus
      Exit Sub
   End If

   If Val(La_nPCs) > Val(Tx_nLic) Then
      MsgBox1 "La cantidad de equipos debe ser menor o igual que la cantidad de licencias.", vbExclamation
      Call Cb_Nivel.SetFocus
      Exit Sub
   End If

   If Ck_Lic.Value = 0 Then
      MsgBox1 "Antes de solicitar su Código de Red debe aceptar las Condiciones de la Licencia de Uso.", vbExclamation
      Call Ck_Lic.SetFocus
      Exit Sub
   End If

   MousePointer = vbHourglass
   DoEvents

   Buf = GenBody()
   If Buf <> "" Then
   
      Buf = "Subject=Solicitud de Código de Red&Body=" & Buf
      
      Buf = ReplaceStr(Buf, vbCr, "%0D")
      Buf = ReplaceStr(Buf, vbLf, "%0A")
      
      Buf = "mailto:" & Trim(gAppCode.emailSop) & "?" & Buf
      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al tratar de enviar el mensaje." & vbLf & "Use el botón copiar para guardar la información y luego péguela en un correo a " & gAppCode.emailSop & ".", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub Cb_Nivel_Click()
   La_Niv = CbItemData(Cb_Nivel)
End Sub

Private Sub Form_Load()

   Call SetupForm

   Call LoadAll

   Bt_Close.Left = Bt_OK.Left
   Bt_Close.Top = Bt_OK.Top
   
   lModif = False
   
   Tx_Cod.Visible = W.InDesign
   La_Niv.Visible = W.InDesign
   
   La_Ver = FwGetVer()
   
End Sub

Private Sub LoadAll()
   Dim r As Integer, Buf As String, nLic As Integer, Aut As Boolean, i As Integer, bChk As Boolean
   Dim MyCod As String, Dt As Long, c As Integer
   
   If gOficina.Rut = "" Then
      MsgBox1 "Debe ingresar el RUT de su empresa en el menú Configuración >> Datos Oficina.", vbExclamation
   End If
   
   Tx_RUT = FmtRut(vFmtRut(gOficina.Rut))
   Tx_Oficina = gOficina.Nombre
   
   La_NetCode.Visible = lAdm
   Tx_NetCode.Visible = lAdm
   Bt_OK.Visible = lAdm
   Bt_Cancel.Visible = lAdm
   
   Bt_Close.Visible = Not lAdm
   Bt_DelPC.Visible = Not lAdm
   Bt_Copy.Visible = Not lAdm
   'Bt_Send.Visible = Not lAdm
   Ck_Lic.Visible = Not lAdm
   La_InfoLic.Visible = Not lAdm
   ' Bt_Lic.Visible = Not lAdm
   
   If lAdm Then
      Me.Caption = "Ingreso de Código de Licencia de Red"
   Else
      Me.Caption = "Solicitud de Código de Licencia de Red"
   End If
   
   For i = 0 To UBound(gAppCode.Nivel)
      If gAppCode.Nivel(i).Desc = "" Then
         Exit For
      End If
      Call AddItem(Cb_Nivel, gAppCode.Nivel(i).Desc, gAppCode.Nivel(i).id)
   Next i

   Cb_Nivel.ListIndex = 0

   Call SelItem(Cb_Nivel, Val(FwDecrypt1(GetIniString(gLicFile, PC_INFO, PC_NIV & 3), KEY_CRYP + 3147)) - 654321)
   'Tx_RUT = FwDecrypt1(GetIniString(gLicFile, PC_INFO, PC_RUT & 2), KEY_CRYP + 5217)
      
   MyCod = FwGetPcCode()
      
   bChk = CheckInscPC()
   Buf = ""
   nLic = 0
   r = 0
   
   For i = 1 To MaxLic

      If GetIniString(gLicFile, PC_EQUIP, PC_NOM & i) <> "" Then
         r = r + 1
         
         If r >= Grid.rows Then
            Grid.rows = r + 1
         End If
         
         Grid.TextMatrix(r, C_IDX) = i
'         Grid.TextMatrix(r, C_PC) = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_NOM & i), KEY_CRYP + i * 10)
         Grid.TextMatrix(r, C_PC) = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_NOM & i), KEY_CRYP + i * 10)
         Grid.TextMatrix(r, C_MAC) = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_MAC & i), KEY_CRYP + i * 30)
         Grid.TextMatrix(r, C_CODPC) = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_COD & i), KEY_CRYP + i * 75)
         Grid.TextMatrix(r, C_CLIC) = FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_AUT & i), KEY_CRYP + i * 155)
         
         Dt = Val(FwDecrypt1(GetIniString(gLicFile, PC_EQUIP, PC_ULT & i), KEY_CRYP + i * 137))
         If Dt > 0 Then
            Grid.TextMatrix(r, C_ULT) = FmtFecha(Dt)
         End If
         
         If Grid.TextMatrix(r, C_CLIC) = "Si" And bChk = True Then
            nLic = nLic + 1
         Else
            Grid.TextMatrix(r, C_CLIC) = "No"
         End If
         
         If StrComp(W.PcName, Grid.TextMatrix(r, C_PC), vbTextCompare) = 0 And StrComp(MyCod, Grid.TextMatrix(r, C_CODPC), vbTextCompare) = 0 Then
            Call FGrForeColor(Grid, r, C_PC, vbBlue)
            Call FGrForeColor(Grid, r, C_CODPC, vbBlue)
            
            If Grid.TextMatrix(r, C_MAC) = W.Mac Then
               Call FGrForeColor(Grid, r, C_MAC, vbBlue)
            End If
         End If
         
      End If
   Next i
            
   Call FGrVRows(Grid)
   
   Buf = GetIniString(gLicFile, PC_INFO, PC_NLIC & 3) ' si está el dato, lo mostramos
   If Buf <> "" Then
      nLic = (Val(FwDecrypt1(Buf, KEY_CRYP + 5043)) - 735081) / 19
      
      If bChk Then
         La_Aut = nLic & " licencias autorizadas"
      Else
         La_Aut = "No hay licencias autorizadas"
      End If
      La_Aut.Visible = True
   Else
      La_Aut.Visible = False
   End If
   
End Sub

Private Sub SaveAll()
   Dim Q1 As String, Rc As Long
   Dim r As Integer, Buf As String, i As Integer, nPCs As Integer
   
   nPCs = 0
   For r = 1 To Grid.rows - 1
   
      If Grid.TextMatrix(r, C_PC) <> "" Then
      
         i = Val(Grid.TextMatrix(r, C_IDX))
         Buf = Grid.TextMatrix(r, C_LIC)
         
         If StrComp(Buf, "Si", vbTextCompare) = 0 Then
            nPCs = nPCs + 1
         Else
            Buf = "N" & r
         End If
         
         Call SetIniString(gLicFile, PC_EQUIP, PC_AUT & i, FwEncrypt1(Buf, KEY_CRYP + i * 155))
      
      End If
   Next r

   If nPCs > 0 Then
      Buf = 654321 + CbItemData(Cb_Nivel)
      Call SetIniString(gLicFile, PC_INFO, PC_NIV & 3, FwEncrypt1(Buf, KEY_CRYP + 3147))
      Buf = 735081 + 19 * Val(Tx_nLic)
      Call SetIniString(gLicFile, PC_INFO, PC_NLIC & 3, FwEncrypt1(Buf, KEY_CRYP + 5043))
      Call SetIniString(gLicFile, PC_INFO, PC_NCOD & 1, FwEncrypt1(Tx_NetCode, KEY_CRYP + 2345))
      Call SetIniString(gLicFile, PC_INFO, PC_RUT & 1, FwEncrypt1(Tx_RUT, KEY_CRYP + 7145))
   End If

End Sub

Private Sub Grid_DblClick()
   Dim Row As Integer, Col As Integer
   Dim Value As Integer
   
   Row = Grid.Row
   Col = Grid.Col

   If Col <> C_LIC Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_PC) = "" Then
      Exit Sub
   End If
   
   lModif = True
      
   If UCase(Grid.TextMatrix(Row, C_LIC)) <> "SI" Then
      Grid.TextMatrix(Row, C_LIC) = "Si"
   Else
      Grid.TextMatrix(Row, C_LIC) = "No"
   End If
   
   Call GenPcInfo
   Tx_nLic = La_nPCs
   
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, "Equipos Inscritos")
   End If
   
End Sub
Private Function GenInfo1() As String
   Dim Info As String, Buf As String
   Dim i As Integer
   
   Info = GenInfo()

   If Info <> "" Then
   
'      Info = Info & "RUT:" & Tx_RUT & ";" & vbCrLf
'      Info = Info & "PRD:" & APP_NAME & ";" & vbCrLf
   
      Buf = ReplaceStr(Info, " ", "")
      Buf = ReplaceStr(Buf, vbCr, "")
      Buf = ReplaceStr(Buf, vbLf, "")
      
      Info = Info & "Verif: " & (GenClave(Buf, 54321) Mod 65432) & ";" & vbCrLf

   End If
   
   GenInfo1 = Info
   
End Function


Private Function GenPcInfo() As String
   Dim r As Integer, Info As String
   
   Info = ""
   lnPCs = 0
   For r = 1 To Grid.rows - 1
      
      If UCase(Grid.TextMatrix(r, C_LIC)) = "SI" Then
         Info = Info & "::" & Grid.TextMatrix(r, C_PC) & ":" & Grid.TextMatrix(r, C_MAC) & ":" & Grid.TextMatrix(r, C_CODPC) & ":" & vbCrLf
         lnPCs = lnPCs + 1
      End If
      
   Next r

   La_nPCs = lnPCs
      
   GenPcInfo = Info

End Function
Private Function GenInfo() As String
   Dim r As Integer, Info As String, nLic As Integer
   
   Info = GenPcInfo()

   If Info = "" Then
      GenInfo = ""
      MsgBox1 "No hay equipos marcados 'Licencia'." & vbCrLf & "Marque los equipos que podrán usar el sistema con el nuevo código.", vbExclamation
      Exit Function
   End If
      
   nLic = Val(Tx_nLic)
      
   Info = Info & "[" & CbItemData(Cb_Nivel) & ":" & lnPCs & ":" & nLic & "]" & vbCrLf

   Info = Info & "RUT:" & Tx_RUT & ";" & vbCrLf
   Info = Info & "PRD:" & APP_NAME & ";" & vbCrLf

   GenInfo = Info

End Function

Public Sub Solicitud()
   lAdm = False
   
   Me.Show vbModal
   
End Sub

Public Sub Admin()
   lAdm = True
   
   Me.Show vbModal
   
End Sub

Private Sub Tx_NetCode_KeyPress(KeyAscii As Integer)
   Call KeyUpper(KeyAscii)
End Sub

Private Sub Tx_Rut_GotFocus()
   Dim r As Long
   
   If Trim(Tx_RUT) <> "" Then
      r = vFmtRut(Tx_RUT)
      If r > 0 Then
         Tx_RUT = r & "-" & DV_Rut(r)
      End If
   End If

End Sub

Private Sub Tx_Rut_LostFocus()
   Dim r As Long

   If Trim(Tx_RUT) <> "" Then
      r = vFmtRut(Tx_RUT)
      If r = 0 Then
         MsgBox1 "El RUT ingresado es inválido.", vbExclamation
      Else
         Tx_RUT = FmtRut(r)
      End If
   End If

End Sub

Private Function GenBody() As String
   Dim Info As String, Buf As String

   Info = GenInfo1()
   
   If Info <> "" Then

      Buf = "Por favor complete los siguientes datos antes de enviar la solicitud." & vbCrLf & vbCrLf
      Buf = Buf & "Empresa: " & gOficina.Nombre & vbCrLf
      Buf = Buf & "RUT Empresa: " & Tx_RUT & vbCrLf
      Buf = Buf & "Nombre del solicitante: " & vbCrLf
      Buf = Buf & "Teléfono: " & vbCrLf
      Buf = Buf & "email: " & vbCrLf
      Buf = Buf & vbCrLf
      Buf = Buf & "*** NO MODIFIQUE ESTA INFORMACION ***" & vbCrLf
      Buf = Buf & "Fecha solicitud: " & Format(Now, "d mmm yyyy") & vbCrLf
      Buf = Buf & "Producto: " & App.Title & vbCrLf
      Buf = Buf & "Versión: " & W.Version & " - " & Format(W.FVersion, "d mmm yyyy") & vbCrLf
      Buf = Buf & "Nivel: [" & CbItemData(Cb_Nivel) & "] " & Cb_Nivel & vbCrLf
      Buf = Buf & "Licencia para " & lnPCs & " equipos" & vbCrLf
      Buf = Buf & ">>>" & vbCrLf & Info
      Buf = Buf & "Ruta: " & W.AppPath & "\" & App.EXEName & ".exe" & vbCrLf
      Buf = Buf & "*************************************" & vbCrLf
      
   End If

   GenBody = Buf
End Function

Private Function NewGenBody() As String
   Dim Info As String, Buf As String
   Dim Page As String
   Dim XmlBody As String
   Dim GetCod As String
   Dim Url As String
   Dim Resp As String
   Info = GenInfo1()
   
   If Info <> "" Then

'      Buf = "Por favor complete los siguientes datos antes de enviar la solicitud." & vbCrLf & vbCrLf
'      Buf = Buf & "Empresa: " & gOficina.Nombre & vbCrLf
'      Buf = Buf & "RUT Empresa: " & Tx_RUT & vbCrLf
'      Buf = Buf & "Nombre del solicitante: " & vbCrLf
'      Buf = Buf & "Teléfono: " & vbCrLf
'      Buf = Buf & "email: " & vbCrLf
'      Buf = Buf & vbCrLf
'      Buf = Buf & "*** NO MODIFIQUE ESTA INFORMACION ***" & vbCrLf
'      Buf = Buf & "Fecha solicitud: " & Format(Now, "d mmm yyyy") & vbCrLf
'      Buf = Buf & "Producto: " & App.Title & vbCrLf
'      Buf = Buf & "Versión: " & W.Version & " - " & Format(W.FVersion, "d mmm yyyy") & vbCrLf
'      Buf = Buf & "Nivel: [" & CbItemData(Cb_Nivel) & "] " & Cb_Nivel & vbCrLf
'      Buf = Buf & "Licencia para " & lnPCs & " equipos" & vbCrLf
'      Buf = Buf & ">>>" & vbCrLf & Info
'      Buf = Buf & "Ruta: " & W.AppPath & "\" & App.EXEName & ".exe" & vbCrLf
'      Buf = Buf & "*************************************" & vbCrLf
      
      'Buf = "Por favor complete los siguientes datos antes de enviar la solicitud." & vbCrLf & vbCrLf
      'Buf = "Empresa: " & gOficina.Nombre & vbCrLf
      Buf = "Empresa: " & vbCrLf
      Buf = Buf & "RUT Empresa: " & Tx_RUT & vbCrLf
      Buf = Buf & "Nombre del solicitante: " & vbCrLf
      Buf = Buf & "Teléfono: " & vbCrLf
      Buf = Buf & "email: " & vbCrLf
      Buf = Buf & vbCrLf
      Buf = Buf & "*** NO MODIFIQUE ESTA INFORMACION ***" & vbCrLf
      Buf = Buf & "Fecha solicitud: " & Format(Now, "d mmm yyyy") & vbCrLf
      Buf = Buf & "Producto: " & App.Title & vbCrLf
      Buf = Buf & "Versión: " & W.Version & " - " & Format(W.FVersion, "d mmm yyyy") & vbCrLf
      Buf = Buf & "Nivel: [" & CbItemData(Cb_Nivel) & "] " & Cb_Nivel & vbCrLf
      Buf = Buf & "Licencia para " & lnPCs & " equipos" & vbCrLf
      Buf = Buf & ">>>" & vbCrLf & Info
      Buf = Buf & "Ruta: " & W.AppPath & "\" & App.EXEName & ".exe" & vbCrLf
      Buf = Buf & "*************************************" & vbCrLf
      
      Call AddLog("Informacion del BUF: " & Buf)
      Url = "https://servicioslp.thomsonreuters.cl/wsActivacion/wsActivacion.asmx"
      XmlBody = BuildReqActivaLicencia("", Buf)
      
      Page = PostWebservice(Url, "http://tempuri.org/ActivacionLicencia", XmlBody)
      
      'Page = FwWebReadPagePost("servicioslp.thomsonreuters.cl", "/WSActivacion/Api/Licencia?Info=" & Buf, 5000)
      Call AddLog("Informacion del servicio: " & Page)
      
      Resp = FwGetXmlTag(Page, "Codigo")
      
      
'      Page = Replace(Replace(Replace(Replace(Page, Chr(34), ""), "{", ""), "}", ""), ":", ",")
'      If Trim(Page) <> "" Then
'      GetCod = Split(Page, ",")(1)
      
        If Resp = "0" Then
             Tx_NetCode = FwGetXmlTag(Page, "Valor") 'Split(Page, ",")(5)
             Call bt_OK_Click
        Else
              MsgBox FwGetXmlTag(Page, "descripcion"), vbExclamation, "Licencias"
        End If
'      End If
      
      
   End If

   NewGenBody = Buf
End Function
