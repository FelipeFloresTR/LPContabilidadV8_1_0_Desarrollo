VERSION 5.00
Begin VB.Form FrmImpActFijos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traer Activos Fijos desde Año Anterior"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8715
   Icon            =   "FrmImpActFijos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   6675
      Begin VB.OptionButton Op_ImpTodos 
         Caption         =   "Traer TODOS  los activos fijos, incluyendo aquellos que han sido traídos anteriormente"
         Height          =   555
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   6075
      End
      Begin VB.OptionButton Op_ImpNuevos 
         Caption         =   "Traer los activos fijos nuevos, que no han sido traídos anteriormente y actualizar los ya traídos."
         Height          =   555
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   6195
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   360
      Picture         =   "FrmImpActFijos.frx":000C
      ScaleHeight     =   780
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   360
      Width           =   750
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5520
      TabIndex        =   0
      Top             =   2280
      Width           =   1155
   End
End
Attribute VB_Name = "FrmImpActFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()

   If Op_ImpTodos Then
      If MsgBox1("ATENCIÓN:" & vbCrLf & vbCrLf & "Esta opción importará nuevamente TODOS los activos fijos desde el año anterior." & vbCrLf & vbCrLf & "Es posible que algunos activos fijos queden duplicados." & vbCrLf & vbCrLf & "Si esto es así, elimine todos los activos fijos duplicados y vuelva a realizar la importación." & vbCrLf & vbCrLf & "NOTA: Si sólo desea traer los nuevos activos fijos y actualizar los ya importados, utilice la otra opción." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      Me.MousePointer = vbHourglass
      Call GenActFijoResidual(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, True)
      Me.MousePointer = vbDefault
      
   Else
      Me.MousePointer = vbHourglass
      Call GenActFijoResidual(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, False)
      Me.MousePointer = vbDefault
   
   End If
   
   Unload Me
   
End Sub
