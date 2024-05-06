VERSION 5.00
Begin VB.Form FrmSelRepActFijo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Control de Activo Fijo"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "FrmSelRepActFijo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   480
      Picture         =   "FrmSelRepActFijo.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   5
      Top             =   660
      Width           =   885
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Index           =   0
      Left            =   1740
      TabIndex        =   4
      Top             =   540
      Width           =   3675
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Control de Activo Fijo Tributario"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   1
         Top             =   780
         Width           =   3075
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Control de Activo Fijo Financiero (IFRS)"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "FrmSelRepActFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INFO_FINANCIERO = 1
Const INFO_TRIBUTARIO = 2

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As Form
   
   
   Me.MousePointer = vbHourglass
   If Op_Informes(INFO_FINANCIERO).Value <> 0 Then
      Set Frm = New FrmRepActFijoIFRS
      Call Frm.FView
   
   Else
      If gMaxCred33 < 0 Then    'el usuario no ha ingresado el Max Cred 33 bis
         If gMaxUTMCred33_Pesos = 0 Then
            If MsgBox1("No se ha ingresado el valor de la UTM. Este valor se utiliza para calcular el máximo para Crédito Art. 33 bis", vbExclamation + vbOKCancel) = vbCancel Then
               Exit Sub
            End If
         Else
            If MsgBox1("Revise si el último valor de la UTM y del IPC ingresados en el sistema están actualizados.", vbInformation + vbOKCancel) = vbCancel Then
               Exit Sub
            End If
            If MsgBox1("Verifique la correcta aplicación del porcentaje del Crédito por Activo Fijo, según instrucciones del Artículo 33 bis Ley de Renta." & vbCrLf & vbCrLf & "Para esto, ingrese a la Configuración Inicial, botón Configurar Impuestos (Menú Configuración).", vbInformation + vbOKCancel) = vbCancel Then
               Exit Sub
            End If
         End If
      End If
      
      Set Frm = New FrmRepActivoFijo
      Call Frm.FView
   End If
   
   Set Frm = Nothing
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Op_Informes(INFO_FINANCIERO) = True
   
End Sub

Private Sub Op_Informes_DblClick(Index As Integer)

   Call PostClick(Bt_Sel)
   
End Sub
