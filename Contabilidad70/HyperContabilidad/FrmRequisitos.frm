VERSION 5.00
Begin VB.Form FrmRequisitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisitos"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   Icon            =   "FrmRequisitos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Req 
      Caption         =   "Depreciación Ley 21.256 Art. 3 (22 Bis TTO Ley 21.210)"
      Height          =   1395
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "3° Fecha Adquisición 01/06/2020 al 31/12/2022"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1° Bienes Fisicos del activo inmovilizado que sean Depreciables"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   4515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2° Bienes Nuevo o Importados"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   660
         Width           =   2160
      End
   End
   Begin VB.Frame Fr_Req 
      Caption         =   "Depreciaciones Régimen Ley 21.210 Art. 21 y 22 Transitorios  y modificaciones Ley  21.256 Art. 3"
      Height          =   4215
      Index           =   5
      Left            =   180
      TabIndex        =   14
      Top             =   5340
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a. Bienes Físicos del activo inmovilizado que sean Depreciables "
         Height          =   195
         Index           =   10
         Left            =   540
         TabIndex        =   24
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "d. Los bienes deben ser instalados físicamente y utilizados exclusivamente en Región de la Araucanía"
         Height          =   195
         Index           =   9
         Left            =   540
         TabIndex        =   23
         Top             =   3720
         Width           =   7230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "c. Fecha Adquisición entre el 01/10/2019 al 31/05/2020 "
         Height          =   195
         Index           =   8
         Left            =   540
         TabIndex        =   22
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "b. Bienes Nuevos o Importados "
         Height          =   255
         Index           =   7
         Left            =   540
         TabIndex        =   21
         Top             =   3000
         Width           =   4575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Depreciación Araucanía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   2070
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "d. Bienes destinados a nuevos proyectos de Inversión"
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   19
         Top             =   1800
         Width           =   3825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "c. Fecha Adquisición entre el 01/10/2019 al 31/05/2020 "
         Height          =   195
         Index           =   5
         Left            =   540
         TabIndex        =   18
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "b. Bienes Nuevos o Importados "
         Height          =   255
         Index           =   3
         Left            =   540
         TabIndex        =   17
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Depreciación Instantánea e Inmediata:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a. Bienes Físicos del activo inmovilizado que sean Depreciables "
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   15
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.Frame Fr_Req 
      Caption         =   "Depreciación Art. 31, 5 bis inc. 1 1/10  (a partir del 1 enero 2020)"
      Height          =   1095
      Index           =   4
      Left            =   180
      TabIndex        =   11
      Top             =   4020
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2° El Bien debe ser Nuevo o Importado"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   2760
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1° Ingresos del Giro es superior a 25.000 UF y no supera las 100.000 UF en los 3 ejercicios anteriores"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   7155
      End
   End
   Begin VB.Frame Fr_Req 
      Caption         =   "Crédito Activo Fijo (Art. 33 Bis Ley de Renta)"
      Height          =   1095
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Lb_Configurar 
         AutoSize        =   -1  'True
         Caption         =   "Configurar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7020
         TabIndex        =   10
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Tx_requisitosCred33bis 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de tasas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5640
         TabIndex        =   9
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2° El rango de ventas promedio de los últimos 3 períodos determina el porcentaje  de crédito a aprovechar"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Width           =   7500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1° Renta efectiva, balance general y contabilidad completa"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4185
      End
   End
   Begin VB.Frame Fr_Req 
      Caption         =   "Depreciación Décima Parte Art. 31, 5 bis inc. 2°"
      Height          =   1095
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1° Ingresos del Giro es superior a 25.000 UF y no supera las 100.000 UF en los 3 ejercicios anteriores"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   7155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2° El Bien debe ser Nuevo o Importado"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   660
         Width           =   2760
      End
   End
   Begin VB.Frame Fr_Req 
      Caption         =   "Depreciación Instantanea Art. 31, 5 bis inc. 1°"
      Height          =   1095
      Index           =   2
      Left            =   180
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2° El Bien debe ser Nuevo o Usado"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   660
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1° Ingresos del Giro es inferior a 25.000 UF en los 3 ejercicios anteriores"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   5070
      End
   End
End
Attribute VB_Name = "FrmRequisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CREDART33 = 1
Const C_DEPINSTANT = 2
Const C_DEPDECIMAPARTE = 3
Const C_DEPDECIMAPARTE2 = 4
Const C_DEPLEY21210 = 5
Const C_DEPLEY21256 = 6

Const NFRAMES = C_DEPLEY21256

Dim lCurFrame As Integer

Private Sub Form_Load()
      
   Fr_Req(lCurFrame).Visible = True
   Fr_Req(lCurFrame).Left = 180
   Fr_Req(lCurFrame).Top = 180
   Me.Height = Fr_Req(lCurFrame).Height + W.YCaption + Fr_Req(lCurFrame).Top + 300
   
   Select Case lCurFrame
      Case C_DEPINSTANT
         Me.Caption = Me.Caption & " Depreciación Instantanea Art. 31, 5 bis inc. 1°"
      Case C_DEPDECIMAPARTE
         Me.Caption = Me.Caption & " Depreciación Décima Parte Art. 31, 5 bis inc. 2°"
      Case C_CREDART33
         Me.Caption = Me.Caption & " Crédito Activo Fijo (Art. 33 Bis Ley de Renta)"
      Case C_DEPDECIMAPARTE2
         Me.Caption = Me.Caption & " Depreciación Art. 31, 5 bis inc. 1 1/10"
      Case C_DEPLEY21210
         Me.Caption = Me.Caption & " Depreciación Regimen Ley 21.210"
      Case C_DEPLEY21256
         Me.Caption = Me.Caption & " Depreciación Regimen Ley 21.256"
   End Select
   
End Sub

Public Function FViewCredArt33bis()
   lCurFrame = C_CREDART33
   Me.Show vbModal
End Function

Public Function FViewDepInstant()
   lCurFrame = C_DEPINSTANT
   Me.Show vbModal
End Function

Public Function FViewDecimaParte()
   lCurFrame = C_DEPDECIMAPARTE
   Me.Show vbModal
End Function

Public Function FViewDecimaParte2()
   lCurFrame = C_DEPDECIMAPARTE2
   Me.Show vbModal
End Function

Public Function FViewLey21210()

   lCurFrame = C_DEPLEY21210
   Me.Show vbModal

End Function
Public Function FViewLey21256()

   lCurFrame = C_DEPLEY21256
   Me.Show vbModal

End Function

Private Sub Lb_Configurar_Click()
   Dim Frm As FrmIVA
   
   Set Frm = New FrmIVA
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub Tx_requisitosCred33bis_Click()
   Dim Frm As FrmHelpCred33bis
   
   Set Frm = New FrmHelpCred33bis
   Frm.Show vbModal
   Set Frm = Nothing

End Sub
