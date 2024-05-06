VERSION 5.00
Begin VB.Form FrmPrtSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar papel para impresora"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "FrmPrtSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   1440
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
      Begin VB.CheckBox Ch_InfoPreliminar 
         Caption         =   "Nota"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   675
      End
      Begin VB.CheckBox Ch_PapelFoliado 
         Caption         =   "Papel Foliado"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informe Preliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   9
         Top             =   700
         Width           =   1245
      End
   End
   Begin VB.CommandButton Bt_ConfigPrt 
      Caption         =   "Configurar Impresora..."
      Height          =   1065
      Left            =   3960
      Picture         =   "FrmPrtSetup.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   540
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orientación Papel"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   2295
      Begin VB.OptionButton Op_Orientacion 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   1
         Top             =   1020
         Width           =   1095
      End
      Begin VB.OptionButton Op_Orientacion 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   0
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   240
         Picture         =   "FrmPrtSetup.frx":0497
         Top             =   900
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "FrmPrtSetup.frx":07A1
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   360
      Picture         =   "FrmPrtSetup.frx":0AAB
      Top             =   480
      Width           =   750
   End
End
Attribute VB_Name = "FrmPrtSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdxOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lRc As Integer

Private Sub bt_Cerrar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_ConfigPrt_Click()

   FrmMain.Cm_PrtDlg.Orientation = lIdxOrientacion
   
   Call PrepararPrt(FrmMain.Cm_PrtDlg)
   
   Op_Orientacion(FrmMain.Cm_PrtDlg.Orientation) = True

End Sub

Private Sub bt_OK_Click()

   lPapelFoliado = (Ch_PapelFoliado.Value <> 0)
   lInfoPreliminar = (Ch_InfoPreliminar.Value <> 0)
   
   lRc = vbOK
   Unload Me
   
End Sub

Private Sub Ch_PapelFoliado_Click()

'   If Not lPapelFoliado Then
'      If Ch_PapelFoliado <> 0 Then
'         MsgBox1 "ADVERTENCIA: Está imprimiendo un Libro No Oficial en Papel Foliado." & vbCrLf & vbCrLf & "Verifique las opciones antes de continuar.", vbInformation
'      End If
'   ElseIf Ch_PapelFoliado = 0 Then
'      MsgBox1 "ADVERTENCIA: Está imprimiendo un Libro Oficial en Papel No Foliado." & vbCrLf & vbCrLf & "Verifique las opciones antes de continuar.", vbInformation
'   End If
      
End Sub

Private Sub Form_Load()
   Op_Orientacion(lIdxOrientacion) = True
   Ch_PapelFoliado = Abs(lPapelFoliado = True)
   
   Call SetupPriv
   
End Sub
Private Sub Op_Orientacion_Click(Index As Integer)
    lIdxOrientacion = Index
End Sub
Public Function FEdit(Orientacion As Integer, PapelFoliado As Boolean, InfoPreliminar As Boolean, Optional ByVal ConfigPrt As Boolean = True) As Integer
   
   lIdxOrientacion = Orientacion
   lPapelFoliado = PapelFoliado
   
   Me.Show vbModal
   
   FEdit = lRc
   Orientacion = lIdxOrientacion
   PapelFoliado = lPapelFoliado
   InfoPreliminar = lInfoPreliminar
   
End Function
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_PapelFoliado = False
      Ch_PapelFoliado.Enabled = False
   End If
   
End Function
