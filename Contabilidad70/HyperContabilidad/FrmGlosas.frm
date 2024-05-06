VERSION 5.00
Begin VB.Form FrmGlosas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Glosas Predefinidas"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "FrmGlosas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   8400
      TabIndex        =   6
      Top             =   960
      Width           =   1095
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmGlosas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nueva cuenta"
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "&Seleccionar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmGlosas.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmGlosas.frx":0A6C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminar glosa seleccionada"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Edit 
         Caption         =   "Edi&tar"
         Height          =   855
         Left            =   0
         Picture         =   "FrmGlosas.frx":10CE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modificar glosa seleccionada"
         Top             =   1980
         Width           =   1095
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8400
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosas"
      Height          =   4935
      Left            =   300
      TabIndex        =   7
      Top             =   300
      Width           =   7815
      Begin VB.ListBox Ls_Glosas 
         Height          =   4545
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7575
      End
   End
End
Attribute VB_Name = "FrmGlosas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lGlosa As String
Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Q1 As String
   
   If Ls_Glosas.ListIndex < 0 Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
      
   If MsgBox1("¿Está seguro que desea eliminar la glosa?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
      MousePointer = vbDefault
      Exit Sub
   End If
            
'   Call ExecSQL(DbMain, "DELETE FROM Glosas WHERE idGlosa = " & ItemData(Ls_Glosas))
   
   Q1 = " WHERE idGlosa = " & ItemData(Ls_Glosas)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call DeleteSQL(DbMain, "Glosas", Q1)
   
   Ls_Glosas.RemoveItem Ls_Glosas.ListIndex
   If Ls_Glosas.ListCount > 0 Then
      Ls_Glosas.ListIndex = 0
   End If
   
   MousePointer = vbDefault
End Sub

Private Sub Bt_Edit_Click()
   Dim Frm As FrmGlosasUpdate
   Dim idGlosa As Long
   
   If Ls_Glosas.ListIndex < 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmGlosasUpdate
   lGlosa = Ls_Glosas
   idGlosa = ItemData(Ls_Glosas)
   
   If Frm.FEdit(lGlosa, idGlosa, O_EDIT) = vbOK Then
      'Modifico la lista
      Ls_Glosas.List(Ls_Glosas.ListIndex) = lGlosa
      
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmGlosasUpdate
   Dim idGlosa As Long
   
   Set Frm = New FrmGlosasUpdate
   If Frm.FEdit(lGlosa, idGlosa, O_NEW) = vbOK Then
      'Agrego a la lista
      Ls_Glosas.AddItem lGlosa
      Ls_Glosas.ItemData(Ls_Glosas.NewIndex) = idGlosa
      Ls_Glosas.ListIndex = Ls_Glosas.NewIndex
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Sel_Click()
   If Ls_Glosas.ListIndex < 0 Then
      Exit Sub
   End If
   
   lGlosa = Ls_Glosas
   lRc = vbOK
   Unload Me
End Sub

Public Function FSelect(ByVal NewGlosa As String) As String
   
   lGlosa = NewGlosa

   Me.Show vbModal
   
   If lRc = vbOK Then
      FSelect = lGlosa
   Else
      FSelect = ""
   End If
   
End Function

Public Sub FEdit()
   
   'permite modificar lista de glosas (New, Mod, Del)
   
   Me.Show vbModal
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   Q1 = "SELECT Glosa,idGlosa FROM Glosas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY Glosa "
   Call FillCombo(Ls_Glosas, DbMain, Q1, -1)
   
End Sub

Private Sub Ls_Glosas_DblClick()
   Call Bt_Sel_Click
End Sub
