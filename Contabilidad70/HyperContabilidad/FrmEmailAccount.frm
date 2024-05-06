VERSION 5.00
Begin VB.Form FrmEmailAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion de Correo saliente"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3975
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      Begin VB.TextBox Tx_body 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   20
         Tag             =   "#PH"
         ToolTipText     =   "Cuenta del usuario de correo"
         Top             =   3000
         Width           =   3855
      End
      Begin VB.CheckBox ck_VerClave 
         Caption         =   "Ver clave"
         Height          =   195
         Left            =   3660
         TabIndex        =   7
         ToolTipText     =   "Permite ver la clave que se está ingresando."
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox Tx_SMTP 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   3
         Tag             =   "#PH"
         ToolTipText     =   "Servidor de correo saliente"
         Top             =   1260
         Width           =   3855
      End
      Begin VB.TextBox Tx_Port 
         Height          =   315
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox Tx_Account 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   5
         Tag             =   "#PH"
         ToolTipText     =   "Cuenta del usuario de correo"
         Top             =   2100
         Width           =   3855
      End
      Begin VB.TextBox Tx_Passw 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1620
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2520
         Width           =   1875
      End
      Begin VB.TextBox Tx_email 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Tx_Asunto 
         Height          =   315
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "En el caso de Gmail debe autorizar el envío desde una aplicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   720
         MouseIcon       =   "FrmEmailAccount.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   21
         ToolTipText     =   "Revise su autorización en GMail"
         Top             =   3600
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje:"
         Height          =   195
         Index           =   10
         Left            =   420
         TabIndex        =   19
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label La_GM 
         AutoSize        =   -1  'True
         Caption         =   "En el caso de Gmail debe autorizar el envío desde una aplicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   420
         MouseIcon       =   "FrmEmailAccount.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   18
         ToolTipText     =   "Revise su autorización en GMail"
         Top             =   5520
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(GMail: 465, Office365: 25, Otros: 25 o 26)"
         Height          =   195
         Index           =   9
         Left            =   2520
         TabIndex        =   17
         Top             =   1740
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servidor SMTP:"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   15
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Puerto:"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   13
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   12
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email para:"
         Height          =   195
         Index           =   5
         Left            =   420
         TabIndex        =   11
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asunto:"
         Height          =   195
         Index           =   6
         Left            =   420
         TabIndex        =   10
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Top             =   960
      Width           =   1275
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Enviar"
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Para Office 365 el servidor SMTP podría ser del tipo: empresa-cl.mail.protection.outlook.com "
      Height          =   1155
      Index           =   7
      Left            =   7320
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   240
      Picture         =   "FrmEmailAccount.frx":02A4
      Top             =   540
      Width           =   825
   End
End
Attribute VB_Name = "FrmEmailAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Private Sub Bt_Cancel_Click()
Unload Me
End Sub

Private Sub Bt_OK_Click()

MousePointer = vbHourglass
   gEmail.smtp = Tx_SMTP
   gEmail.puerto = Tx_Port
   gEmail.Cuenta = Tx_Account
   gEmail.contraseña = Tx_Passw
   gEmail.to = Tx_email
   gEmail.From = Tx_Account
   gEmail.Subject = Tx_Asunto
   gEmail.Body = Tx_body
   

Call SendEmailBalEstados
MousePointer = vbDefault
End Sub



Private Sub ck_VerClave_Click()
If ck_VerClave.Value = 1 Then
Tx_Passw.PasswordChar = ""
Else
Tx_Passw.PasswordChar = "*"
End If
End Sub

Private Sub Form_Load()
   gEmail.smtp = ""
   gEmail.puerto = ""
   gEmail.Cuenta = ""
   gEmail.contraseña = ""
   gEmail.to = ""
   gEmail.From = ""
   gEmail.Subject = ""
   'gEmail.adjunto = ""
   
   'Me.Hide
End Sub

Public Function FEdit(ByVal adjunto As String) As Boolean
   
   FEdit = True
   
   gEmail.adjunto = adjunto
   'Me.Show
End Function

'2861570
Sub SendEmailBalEstados()
On Error Resume Next ' Set up error checking

'Set cdoMsg = Nothing
'Set cdoConf = Nothing
'Set cdoFields = Nothing

Set cdoMsg = CreateObject("CDO.Message")
Set cdoConf = CreateObject("CDO.Configuration")

Set cdoFields = cdoConf.Fields
' Send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
cdoFields.Item(schema & "sendusing") = 2
cdoFields.Item(schema & "smtpserver") = gEmail.smtp
cdoFields.Item(schema & "smtpserverport") = CInt(gEmail.puerto)
cdoFields.Item(schema & "smtpauthenticate") = 1
cdoFields.Item(schema & "sendusername") = gEmail.Cuenta
cdoFields.Item(schema & "sendpassword") = gEmail.contraseña
cdoFields.Item(schema & "smtpusessl") = 1 '".T."
cdoFields.Update
With cdoMsg
    .to = gEmail.to
    .From = gEmail.From
    .Subject = gEmail.Subject
    ' Body of message can be any HTML code
    .HTMLBody = gEmail.Body
    ' Add any attachments to the message
    .AddAttachment gEmail.adjunto
    Set .Configuration = cdoConf
    ' Send the message
    .Send
End With
'Check for errors and display message
If ERR.Number = 0 Then
      MsgBox "Email Enviado Exitosamente", , "Email"
Else
'        ERR.Clear
'        Set cdoMsg = Nothing
'        Set cdoConf = Nothing
'        Set cdoFields = Nothing

        Set cdoMsg = CreateObject("CDO.Message")
        Set cdoConf = CreateObject("CDO.Configuration")

        Set cdoFields = cdoConf.Fields
        ' Send one copy with Google SMTP server (with autentication)
        schema = "http://schemas.microsoft.com/cdo/configuration/"
        cdoFields.Item(schema & "sendusing") = 2
        cdoFields.Item(schema & "smtpserver") = gEmail.smtp
        cdoFields.Item(schema & "smtpserverport") = CInt(gEmail.puerto)
        'cdoFields.Item(schema & "smtpconnectiontimeout") = 60
        cdoFields.Item(schema & "smtpauthenticate") = 1 '".T."
        cdoFields.Item(schema & "sendusername") = gEmail.Cuenta
        cdoFields.Item(schema & "sendpassword") = gEmail.contraseña
        cdoFields.Item(schema & "smtpusessl") = ".T."
       cdoFields.Update
        
        With cdoMsg
            .to = gEmail.to
            .From = gEmail.From
            .Subject = gEmail.Subject
            ' Body of message can be any HTML code
            .HTMLBody = gEmail.Body
            ' Add any attachments to the message
            .AddAttachment gEmail.adjunto
            Set .Configuration = cdoConf
            ' Send the message
            .Send
        End With
        'Check for errors and display message
        If ERR.Number = 0 Or ERR.Number = 3749 Then
              MsgBox "Email Enviado Exitosamente", , "Email"
        Else
              MsgBox "Email Error" & ERR.Description, , "Email"
        End If


     ' MsgBox "Email Error" & ERR.Description, , "Email"
End If
Set cdoMsg = Nothing
Set cdoConf = Nothing
Set cdoFields = Nothing
End Sub

'2861570

Private Sub Label2_Click()
Dim r As Long
   r = ShellExecute(0, "open", "https://support.google.com/a/answer/176600?hl=es#zippy=%2Cutilizar-el-servidor-smtp-de-gmail", 0, 0, 1)
End Sub
