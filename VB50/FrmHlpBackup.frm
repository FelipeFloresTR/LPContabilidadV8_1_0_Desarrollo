VERSION 5.00
Begin VB.Form FrmHlpBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldos"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "FrmHlpBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Msg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmHlpBackup.frx":000C
      Top             =   780
      Width           =   10515
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   9840
      Top             =   180
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   180
      Picture         =   "FrmHlpBackup.frx":0012
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "¿ Respaldó su información esta semana ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "FrmHlpBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Image2 = Image1

   Tx_Msg = "* * *  IMPORTANTE  * * *"
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Es de suma importancia realizar respaldos de la información en forma periódica. Es responsabilidad del usuario o empresa definir una política adecuada al respecto."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "En el caso de la pérdida de información debido al ataque de un virus, la falla de un disco, etc., la única forma de recuperar y no perder el trabajo de meses, es recurrir a los respaldos."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Los programas pueden ser instalados nuevamente, pero si no hay respaldos, la información ingresada se perderá irremediablemente."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Los respaldos deben se hechos en un medio externo, no deben ser hechos en el mismo disco o en el mismo equipo en que se encuentra la aplicación. Un virus puede destruir el contenido de todo el disco o los discos del equipo, o bien puede fallar el disco en que se hizo el respaldo."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Es importante verificar que los respaldos queden bien hechos, de modo que cuando se necesiten puedan ser utilizados. Para esto es bueno probar a recuperar un respaldo y ver si la información es la correcta."
   Tx_Msg = Tx_Msg & " Los dispositivos donde se hace el respaldo (ej CDs), es recomendable que se almacenen fuera de la oficina."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Una forma segura, sencilla y económica es utilizar CDs o DVDs. Estos permiten almacenar gran cantidad de información."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Para nuestra aplicación " & App.Title & ", usted debería respaldar toda la carpeta '" & W.AppPath & "'."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   
   If gDbType = SQL_MYSQL Then
      Tx_Msg = Tx_Msg & "En esta versión la base de datos está en un servidor MySQL, debe solicitar la asistencia de un técnico para realizar el respado de los datos."
      Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   End If
   
   Tx_Msg = Tx_Msg & "Si el respaldo lo hace hoy, cree en el CD una carpeta llamada '" & FmtFecha(Now) & "'. En esta carpeta agregue el contenido de la carpeta '" & W.AppPath & "' y toda otra información importante para usted."
   Tx_Msg = Tx_Msg & vbCrLf
   Tx_Msg = Tx_Msg & "En el siguiente respaldo, utilice otro CD, cree una carpeta con la nueva fecha y agregue en esta nueva carpeta su información."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Si sólo respalda esta aplicación en el CD, seguramente podrá realizar varios respaldos en el mismo CD. Sin embargo, no sería recomendable tener más de cuatro respaldos seguidos en el mismo CD, porque si éste se daña se pierde toda su información."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Se recomienda tener dos o más CDs (CD1, CD2, CD3, ...) e ir usando un CD distinto cada vez, primero el CD1, luego el CD2, después el CD3, ... luego nuevamente el CD1 y así. De ese modo si se daña un CD quedan los otros."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   Tx_Msg = Tx_Msg & "Recuerde mantener actualizado su Antivirus y chequear periódicamente sus discos para reducir los riesgos."
   Tx_Msg = Tx_Msg & vbCrLf & vbCrLf
   
End Sub

