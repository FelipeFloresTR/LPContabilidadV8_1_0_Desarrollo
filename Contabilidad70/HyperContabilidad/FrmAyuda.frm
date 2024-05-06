VERSION 5.00
Begin VB.Form FrmAyuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   Icon            =   "FrmAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Help 
      Caption         =   "Ejemplos de Otros Ajustes (Aumentos)"
      Height          =   2355
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SOLO 14D3       +       Reposición deducción por pago IDPC Voluntario en años anteriores"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   6375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SOLO 14D3       +       Franquicia Letra E, art 14 LIR"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   3675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SOLO 14D3       +        Ingresos exentos de IDPC"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SOLO 14D3       +         Ingresos no rentas generadas por la empresa"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   420
         Width           =   4860
      End
   End
   Begin VB.Frame Fr_Help 
      Caption         =   "Ejemplos de Otros Ajustes (Disminuciones)"
      Height          =   1935
      Index           =   2
      Left            =   180
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SOLO 14D8       -       Incremento asociado a retiros o dividendos recibidos	"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AMBOS             -         Crédito total disponible por IPE	"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   8
         Top             =   900
         Width           =   3765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AMBOS             -         Ingreso diferido imputado en el ejercicio"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   4395
      End
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_OTROSAJUSTAUMENTOS = 1
Const C_OTROSAJUSTDISMIN = 2

Const NFRAMES = C_OTROSAJUSTDISMIN

Dim lCurFrame As Integer
Dim lTitulo As String

Private Sub Form_Load()
      
   Fr_Help(lCurFrame).Visible = True
   Fr_Help(lCurFrame).Top = 180
   Me.Height = Fr_Help(lCurFrame).Height + W.YCaption + Fr_Help(lCurFrame).Top + 300
   
   Me.Caption = Me.Caption & " " & lTitulo
   
End Sub

Public Function FViewOtrosAjustesAumentos(ByVal Titulo As String)
   lTitulo = Titulo
   lCurFrame = C_OTROSAJUSTAUMENTOS
   Me.Show vbModal
End Function

Public Function FViewOtrosAjustesDismin(ByVal Titulo As String)
   lTitulo = Titulo
   lCurFrame = C_OTROSAJUSTDISMIN
   Me.Show vbModal
End Function

