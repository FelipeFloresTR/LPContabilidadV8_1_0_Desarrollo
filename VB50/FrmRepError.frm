VERSION 5.00
Begin VB.Form FrmRepError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Problema"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "FrmRepError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Copy 
      Caption         =   "Copiar"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmRepError.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C00000&
      Height          =   555
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FrmRepError.frx":0456
      Top             =   6660
      Width           =   9015
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7860
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Send 
      Caption         =   "Enviar"
      Height          =   675
      Left            =   7860
      Picture         =   "FrmRepError.frx":04DE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Tx_Info 
      Height          =   3735
      Left            =   180
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2880
      Width           =   9015
   End
   Begin VB.TextBox Tx_Basico 
      Height          =   1995
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   7515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Por favor complete la siguiente información:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   2640
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Información básica:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "FrmRepError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Copy_Click()

   MousePointer = vbHourglass
   DoEvents

   Clipboard.Clear
   Clipboard.SetText GetDatos()

   MsgBox1 "Los datos ya fueron copiados, ahora debe pegarlos en un mail dirigido a " & gAppCode.emailSop & ".", vbInformation

   MousePointer = vbDefault

End Sub

Private Sub Bt_Send_Click()
   Dim Rc As Long
   
   Rc = SendEmail(Me.hWnd, gAppCode.emailSop, "Soporte Fairware", "Reporte de " & App.Title, GetDatos())

End Sub

Private Sub Form_Load()
   Dim Buf As String

   Text1 = ReplaceStr(Text1, "%email%", gAppCode.emailSop)

   Buf = "Fecha reporte: " & Format(Now, "d mmmm yyyy hh:nn") & vbCrLf
   Buf = Buf & vbCrLf
   Buf = Buf & "Producto: " & App.Title & vbCrLf
   
   Buf = Buf & "   Versión: " & W.Version & " - " & Format(W.FVersion, "d mmm yyyy")
#If DATACON = 2 Then
   Buf = Buf & " - SQL Server"
#Else
   Buf = Buf & " - Access"
#End If

   If gAppCode.Demo Then
      Buf = Buf & " - " & gAppCode.txDemo
   End If
   Buf = Buf & vbCrLf
   
   Buf = Buf & "   Cód: " & FwGetPcCode() & vbCrLf
   Buf = Buf & "   Ruta: " & W.AppPath & "\" & App.EXEName & ".exe" & vbCrLf
'   Buf = Buf & "   Ruta datos: " & DbMain.Name & vbCrLf
   Buf = Buf & vbCrLf
   Buf = Buf & "Configuración Equipo:" & vbCrLf
   Buf = Buf & "   PC: " & W.PcName & vbCrLf
   Buf = Buf & "   MAC: " & GetMac() & vbCrLf
   Buf = Buf & "   S.O.: " & GetVersionInfo() & vbCrLf
   Buf = Buf & "   Formato de número: " & Format(12345.6789, "#,##0.0000") & vbCrLf
   Buf = Buf & "   Formato de moneda: " & Format(12345.6789, "$#,##0.0000") & vbCrLf
   Buf = Buf & "   Formato de fecha: " & Format(Now, "mmmm, dddd ") & " " & Now & vbCrLf
   Buf = Buf & "   T/F: " & (1 = 1) & "/" & (1 = 0) & vbCrLf
   Tx_Basico = Buf
   
   Buf = "Identificación del Cliente:" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "Nombre del reportante:" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "Teléfono del reportante:" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "Descripción del problema:" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "¿En que ventana ocurre?" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "¿En que menú ocurre?" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "¿Qué pasos hay que seguir para reproducirlo?" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "¿Qué mensaje de error aparece?" & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Buf = Buf & "Si existe la carpeta " & W.AppPath & "\LOG, por favor adjunte los archivos de los últimos días." & vbCrLf
   Buf = Buf & vbCrLf & vbCrLf
   Tx_Info = Buf
   
End Sub

Private Function GetDatos() As String
   Dim Body As String

   Body = Trim(Tx_Info) & vbCrLf
   Body = Body & String(30, "=") & vbCrLf
   Body = Body & Trim(Tx_Basico) & vbCrLf
 
   GetDatos = Body
   
End Function
