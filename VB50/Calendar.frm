VERSION 5.00
Begin VB.Form FrmCalendar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   ForeColor       =   &H80000008&
   Icon            =   "Calendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   1260
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   41
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   41
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   40
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   40
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   39
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   39
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   38
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   38
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   37
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   37
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   36
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   36
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   35
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   35
      Text            =   "00"
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Bt_Today 
      Caption         =   "&Hoy"
      Height          =   315
      Left            =   2280
      TabIndex        =   43
      Top             =   2220
      Width           =   855
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   34
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   34
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   33
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   33
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   32
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   32
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   31
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   31
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   30
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   30
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   29
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   29
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   28
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   28
      Text            =   "00"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   27
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   27
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   26
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   26
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   25
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   25
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   24
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   24
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   23
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   23
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   22
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   22
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   21
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   21
      Text            =   "00"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   20
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   20
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   19
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   19
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   18
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   18
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   17
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   17
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   16
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   16
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   15
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   15
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   14
      Text            =   "00"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   13
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   13
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   12
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   12
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   11
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   11
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   10
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   10
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   9
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   9
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   8
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Text            =   "00"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   215
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   5
      Left            =   1620
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   3
      Left            =   1020
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   1
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Tx_Dia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   215
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Text            =   "00"
      Top             =   840
      Width           =   255
   End
   Begin VB.HScrollBar Hs_Fecha 
      Height          =   255
      LargeChange     =   12
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Value           =   200
      Width           =   2055
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   2280
      TabIndex        =   45
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   2280
      TabIndex        =   44
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Tx_AnoMes 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Do"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   52
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sá"
      Height          =   255
      Index           =   5
      Left            =   1620
      TabIndex        =   51
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Vi"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   50
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ju"
      Height          =   255
      Index           =   3
      Left            =   1020
      TabIndex        =   49
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mi"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   48
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ma"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   47
      Top             =   540
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lu"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   540
      Width           =   255
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 31 oct 2018: se cambia el uso de .Tag por .DataField por uso del .Tag en PrtForm y PrtForm2

Private lRc       As Integer

Private lNow      As Long

Private SelFecha  As Long
Private lFecha    As Long
Private oFecha    As Long
Private iDia      As Integer
Private lLockMonth As Boolean

Public Function SelDate(Fecha As Long, Optional ByVal AnosAntes As Integer = 70, Optional ByVal AnosDespues As Integer = 20) As Integer
   Dim Dif As Long

   lNow = Int(Now)

   If Fecha = 0 Then
      Fecha = lNow
   End If

   oFecha = Fecha
   lFecha = Fecha
   SelFecha = lFecha

   Hs_Fecha.Min = -12 * AnosAntes
   Hs_Fecha.Max = 12 * AnosDespues
   
   Dif = ((Year(oFecha) - 1950) * 12 + Month(oFecha)) - ((Year(lNow) - 1950) * 12 + Month(lNow))
   
   If Dif < -12 * 70 Or Dif > 12 * 20 Then
      Dif = 0
   End If
   
   Hs_Fecha.Value = Dif
   
   Me.Show vbModal
   
   Fecha = SelFecha
   
   SelDate = lRc
   
End Function

Private Sub FillCal(ByVal Fecha As Long)
   Dim Tm1 As Long, Tm2 As Long, d As Long
   Dim i As Integer
   
   On Error Resume Next
   
   Call FirstLastMonthDay(Fecha, Tm1, Tm2)

   Tx_AnoMes = gNomMes(Month(Tm1)) & " " & Year(Tm1)

   If Tx_AnoMes = gNomMes(Month(lNow)) & " " & Year(lNow) Then
      Tx_AnoMes.ForeColor = vbActiveTitleBar
   Else
      Tx_AnoMes.ForeColor = vbWindowText
   End If
   
   i = 0
   Do While i < Weekday(Tm1, vbMonday)
      Tx_Dia(i).Visible = False
      'Tx_Dia(i) = ""
      'Tx_Dia(i).BorderStyle = 0
      'Tx_Dia(i).BackColor = vbButtonFace
      
      i = i + 1
   Loop

   i = i - 1
   For d = 0 To Tm2 - Tm1
   
      Tx_Dia(i + d).Visible = True
      Tx_Dia(i + d) = d + 1
      Tx_Dia(i + d).BorderStyle = 0
      Tx_Dia(i + d).BackColor = vbButtonFace
      
      If Tm1 + d = Fecha Then
         iDia = i + d
         Tx_Dia(iDia).BorderStyle = 1
         Tx_Dia(iDia).SetFocus
      End If
      
      If Tm1 + d = lNow Then
         If Weekday(lNow) = vbSunday Then
            Tx_Dia(i + d).BackColor = vbRed
         Else
            Tx_Dia(i + d).BackColor = vbActiveTitleBar
         End If
         Tx_Dia(i + d).ForeColor = vbWhite
      Else
         If Weekday(Tm1 + d) = vbSunday Then
            Tx_Dia(i + d).ForeColor = vbRed
         Else
            Tx_Dia(i + d).ForeColor = vbWindowText
         End If
      End If
      
   Next d

   For i = i + d To 41
      Tx_Dia(i).Visible = False

      'Tx_Dia(i) = ""
      'Tx_Dia(i).BorderStyle = 0
      'Tx_Dia(i).BackColor = vbButtonFace
   Next i

End Sub

Private Sub Bt_OK_Click()
      
   SelFecha = lFecha
   lRc = vbOK
   
   Unload Me
   
End Sub

Private Sub Bt_Today_Click()
   lFecha = lNow
   Hs_Fecha.Value = 0
   Call Hs_Fecha_Scroll
End Sub

Private Sub Form_Activate()

   If iDia >= 0 Then
      Call Tx_Dia(iDia).SetFocus
   End If
   
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   
   If lLockMonth Then
      Hs_Fecha.Visible = False
      Bt_Today.Visible = False
      Me.Height = Me.Height - Bt_Today.Height
   End If
   
End Sub

Private Sub Hs_Fecha_Change()
   Call Hs_Fecha_Scroll
End Sub

Private Sub Hs_Fecha_Scroll()
   Dim Dia As Integer, LDia As Integer, LDia1 As Integer
   Dim oFecha As Long
   
   oFecha = lFecha
   Dia = Day(oFecha)
   LDia = Day(DateSerial(Year(oFecha), Month(oFecha) + 1, 1) - 1)
   
   lFecha = DateSerial(Year(lNow), Month(lNow) + Hs_Fecha.Value, 1)
   LDia1 = Day(DateSerial(Year(lFecha), Month(lFecha) + 1, 1) - 1)

   If Dia = LDia Then
      lFecha = lFecha + LDia1 - 1
   ElseIf Dia <= LDia1 Then
      lFecha = lFecha + Dia - 1
   Else
      lFecha = lFecha + LDia1 - 1
   End If
      
   Call FillCal(lFecha)

End Sub

Private Sub Timer1_Timer()
   Static bShow As Boolean

   If bShow = False Then
      bShow = True
      Timer1.Enabled = False
      
      Call AddLog("Calendar: Hs_Fecha.visible=" & Hs_Fecha.Visible & ", bLockMonth=" & lLockMonth & ", x=" & Hs_Fecha.Left & ", y=" & Hs_Fecha.top & ", w=" & Hs_Fecha.Width & ", h=" & Hs_Fecha.Height)
      
      Hs_Fecha.Left = Tx_AnoMes.Left
      Hs_Fecha.top = Bt_Today.top + Bt_Today.Height - Hs_Fecha.Height
      Me.Caption = Me.Caption & "."
      
      If lLockMonth = False Then
         Hs_Fecha.Visible = True
      End If
   End If
End Sub

Private Sub Tx_Dia_DblClick(index As Integer)

   If Tx_Dia(index) <> "" Then
      Call PostClick(bt_Ok)
      'Bt_OK.SetFocus
      'Call SendKeys("{ENTER}", False)
   End If
   
End Sub

Private Sub Tx_Dia_GotFocus(index As Integer)

   If Tx_Dia(index) <> "" Then
      Tx_Dia(iDia).BorderStyle = 0
      iDia = index
      Tx_Dia(iDia).BorderStyle = 1
      lFecha = DateSerial(Year(lFecha), Month(lFecha), Val(Tx_Dia(iDia)))
      
   End If
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Public Function TxSelDate(Tx As TextBox, Optional ByVal bLocate As Boolean = 1, Optional ByVal bLockMonth As Boolean = 0) As Integer
   Dim Dt As Long, bMod As Boolean, oDt As Long

   bMod = Tx.DataChanged
   lLockMonth = bLockMonth

   oDt = GetTxDate(Tx)

   If Trim(Tx) = "" Then
      oDt = 0
   ElseIf Tx.DataField <> "" Then
      oDt = Val(Tx.DataField)
   Else
      oDt = VFmtDate(Tx)
   End If

   Load Me

   If bLocate Then
      Call LocateMe(Tx)

      'lPnt.X = (Tx.Left / Screen.TwipsPerPixelX)
      'lPnt.Y = ((Tx.Top + Tx.Height) / Screen.TwipsPerPixelY) + 1
      'Call ClientToScreen(Tx.Container.hWnd, lPnt)

      'Me.Left = lPnt.X * Screen.TwipsPerPixelX
      'Me.Top = lPnt.Y * Screen.TwipsPerPixelY
      
      'Call FormPos(Me, -1)
            
   End If

   Dt = oDt
   If SelDate(Dt) = vbOK Then
      Tx = Format(Dt, DATEFMT)
      Tx.DataField = Dt
      TxSelDate = vbOK
   Else
      TxSelDate = vbCancel
   End If
   
   Tx.DataChanged = (bMod Or (oDt <> Dt))

End Function

Public Function TxSelMonth(Tx As TextBox, Optional Locate As Boolean = True) As Integer
   Dim Dt As Long
   'Dim lPnt As POINTAPI_T

   If Trim(Tx) = "" Then
      Dt = 0
   Else
      Dt = VFmtDate("01 " & Tx)
   End If

   Load Me

   If Locate Then
      Call LocateMe(Tx)

      'lPnt.X = (Tx.Left / Screen.TwipsPerPixelX)
      'lPnt.Y = ((Tx.Top + Tx.Height) / Screen.TwipsPerPixelY) + 1
      'Call ClientToScreen(Tx.Container.hWnd, lPnt)

      'Me.Left = lPnt.X * Screen.TwipsPerPixelX
      'Me.Top = lPnt.Y * Screen.TwipsPerPixelY
      
      'Call FormPos(Me, -1)
            
   End If

   If SelDate(Dt) = vbOK Then
      Dt = DateSerial(Year(Dt), Month(Dt), 1)
      Tx.DataField = Dt
      Tx = Format(Dt, MONTHFMT)
      TxSelMonth = vbOK
   Else
      TxSelMonth = vbCancel
   End If

End Function

Private Sub Tx_Dia_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      Case vbKeyLeft:
         If index - 1 >= 0 Then
            If Tx_Dia(index - 1).Visible Then
               Tx_Dia(index - 1).SetFocus
            End If
         End If
         
      Case vbKeyUp:
         If index - 7 >= 0 Then
            If Tx_Dia(index - 7).Visible Then
               Tx_Dia(index - 7).SetFocus
            End If
         End If
         
      Case vbKeyRight:
         If index + 1 <= 41 Then
            If Tx_Dia(index + 1).Visible Then
               Tx_Dia(index + 1).SetFocus
            End If
         End If
         
      Case vbKeyDown:
         If index + 7 <= 41 Then
            If Tx_Dia(index + 7).Visible Then
               Tx_Dia(index + 7).SetFocus
            End If
         End If
         
   End Select
   
End Sub

Private Sub LocateMe(Tx As TextBox)
   Dim lPnt As POINTAPI_T

   lPnt.x = ((Tx.Left + Tx.Width - Me.Width + w.xScroll) / Screen.TwipsPerPixelX)
   lPnt.Y = ((Tx.top + Tx.Height) / Screen.TwipsPerPixelY) + 1
   
   Call ClientToScreen(Tx.Container.hWnd, lPnt)

   If lPnt.x < 0 Then
      lPnt.x = 0
   End If
   
   Me.Left = lPnt.x * Screen.TwipsPerPixelX
   Me.top = lPnt.Y * Screen.TwipsPerPixelY
   
   Call FormPos(Me, -1)

End Sub
Public Sub Locate(Frm As Form, ByVal x As Single, Y As Single)
   Dim lPnt As POINTAPI_T

   lPnt.x = (x / Screen.TwipsPerPixelX)
   lPnt.Y = (Y / Screen.TwipsPerPixelY) + 1
   
   Call ClientToScreen(Frm.hWnd, lPnt)

   If lPnt.x < 0 Then
      lPnt.x = 0
   End If
   
   Me.Left = lPnt.x * Screen.TwipsPerPixelX
   Me.top = lPnt.Y * Screen.TwipsPerPixelY
   
   Call FormPos(Me, -1)

End Sub

