VERSION 5.00
Begin VB.Form FrmExpDJAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a HR-Certificados DJ 1924"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   420
      Picture         =   "FrmExpDJAnual.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   480
      Width           =   585
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   2655
      Begin VB.TextBox Tx_Ano 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Exp 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "FrmExpDJAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTipoDJ As String

Public Function ExpHRRAB()

   lTipoDJ = "HRRAB"
   Me.Show vbModal
   
End Function
Public Function ExpHRRAD()

   lTipoDJ = "HRRAD"
   Me.Show vbModal
   
End Function

Public Function Exp1923()

   lTipoDJ = "1923"
   Me.Show vbModal
   
End Function
   
Public Function Exp1924B()

   lTipoDJ = "1924B"
   Me.Show vbModal
   
End Function
Public Function Exp1924C()

   lTipoDJ = "1924C"
   Me.Show vbModal
   
End Function
Public Function Exp1847()

   lTipoDJ = "1847"
   Me.Show vbModal
   
End Function
Public Function ExpRetirosDividendos()

   lTipoDJ = "RetDiv"
   Me.Show vbModal
   
End Function

Private Sub Bt_Close_Click()
   Unload Me
   
End Sub

Private Sub Bt_Exp_Click()
   Dim Frm As FrmDatosDJ1847
   Dim Rc As Integer
   Dim FName As String, FName2 As String

   Select Case lTipoDJ
   
      Case "1923"
         Call Export_DJ1923(FName)
      
      Case "1924B"
         Call Export_DJ1924(lTipoDJ, FName, FName2)
      
      Case "1924C"
         Call Export_DJ1924(lTipoDJ, FName, FName2)
      
      Case "1847"
   
         Set Frm = New FrmDatosDJ1847
         Rc = Frm.FEdit
         Set Frm = Nothing
         
         If Rc = vbOK Then
            MsgBox1 "Recuerde que esta Declaración Jurada podrá ser revisada, completada y/o modificada directamente en la plantilla creada.", vbInformation
            Call Export_DJ1847(FName)
         End If
      
      Case "HRRAB"
         Call Export_DJ1923(FName, True)
         
      Case "RetDiv"
         Call Export_RetirosDividendos(FName)
     
      Case "HRRAD"
         Call Export_HRRAD_BaseImp14D(FName)
         Call Export_HRRAD_CPS_Totales(FName)
         
         '2699582
         If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
         Call Export_PPM_BaseImp14D(FName)
         
         End If
         'fin 2699582
         'Call Export_HRRAD_CPS_Detalle(fname)
      
   End Select
   
   If lTipoDJ <> "HRRAB" And lTipoDJ <> "RetDiv" And lTipoDJ <> "HRRAD" Then
      If gEmpresa.Ano < 2019 Then
         MsgBox1 "Recuerde que debe tomar el archivo csv y llevarlo al HR Importador de Certificados", vbInformation
      ElseIf lTipoDJ = "1924C" Then
         Call ConectHRCertif(lTipoDJ, FName2)
      Else
         Call ConectHRCertif(lTipoDJ, FName)
      End If
   End If
      
End Sub

Private Sub Form_Load()

   Tx_Ano = gEmpresa.Ano
   If lTipoDJ = "HRRAB" Then
      Me.Caption = "Exportar a HR-RAB"
      
      If gEmpresa.Ano >= 2020 Then
         Me.Caption = "Exportar a HR-RAD"
      End If
      
   ElseIf lTipoDJ = "HRRAD" Then
      Me.Caption = "Exportar 14D HR RAD"
   
   ElseIf lTipoDJ = "RetDiv" Then
      If gEmpresa.TipoContrib = CONTRIB_SAABIERTA Or gEmpresa.TipoContrib = CONTRIB_SACERRADA Or gEmpresa.TipoContrib = CONTRIB_SPORACCION Then
         Me.Caption = "Exportar archivo Dividendos"
      Else
         Me.Caption = "Exportar archivo Retiros"
      End If

   Else
      Me.Caption = "Exportar a HR-Certificados DJ " & lTipoDJ
      
   End If
   
   If lTipoDJ = "1923" Then
      Me.Caption = Me.Caption & " (Sección B)"
   End If
   
End Sub

