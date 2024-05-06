VERSION 5.00
Begin VB.Form FrmConfigActFijoIFRS 
   Caption         =   "Configuración de Activos Fijos Financieros"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7740
      TabIndex        =   8
      Top             =   600
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Componentes"
      Height          =   3735
      Left            =   1620
      TabIndex        =   10
      Top             =   1620
      Width           =   5835
      Begin VB.ListBox Ls_Componentes 
         Height          =   2985
         Left            =   300
         TabIndex        =   4
         Top             =   420
         Width           =   3315
      End
      Begin VB.CommandButton Bt_DelComp 
         Caption         =   "&Eliminar"
         Height          =   800
         Left            =   4080
         Picture         =   "FrmConfigActFijoIFRS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar Componente seleccionada"
         Top             =   2100
         Width           =   1095
      End
      Begin VB.CommandButton Bt_EditComp 
         Caption         =   "Edi&tar"
         Height          =   800
         Left            =   4080
         Picture         =   "FrmConfigActFijoIFRS.frx":0662
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modificar Componente seleccionada"
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton Bt_NewComp 
         Caption         =   "&Agregar"
         Height          =   800
         Left            =   4080
         Picture         =   "FrmConfigActFijoIFRS.frx":0C35
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nueva Componente"
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo"
      Height          =   1095
      Left            =   1620
      TabIndex        =   9
      Top             =   480
      Width           =   5835
      Begin VB.ComboBox Cb_Grupo 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3315
      End
      Begin VB.CommandButton Bt_DelGrupo 
         Height          =   480
         Left            =   4980
         Picture         =   "FrmConfigActFijoIFRS.frx":11AD
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar Grupo seleccionado"
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton Bt_EditGrupo 
         Height          =   480
         Left            =   4440
         Picture         =   "FrmConfigActFijoIFRS.frx":180F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modificar Grupo seleccionado"
         Top             =   360
         Width           =   540
      End
      Begin VB.CommandButton Bt_NewGrupo 
         Height          =   480
         Left            =   3840
         Picture         =   "FrmConfigActFijoIFRS.frx":1DE2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo Grupo"
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   300
      Picture         =   "FrmConfigActFijoIFRS.frx":235A
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   5
      Top             =   540
      Width           =   885
   End
End
Attribute VB_Name = "FrmConfigActFijoIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_DelGrupo_Click()
   Dim id As Long
   Dim Nombre As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   id = CbItemData(Cb_Grupo)
   If id <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Q1 = "SELECT Count(*) FROM ActFijoFicha WHERE IdGrupo = " & id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then
         If vFld(Rs(0)) = 1 Then
            MsgBox1 "No es posible eliminar este Grupo. Hay un Activo Fijo que pertenece a éste.", vbExclamation
         Else
            MsgBox1 "No es posible eliminar este Grupo. Hay " & vFld(Rs(0)) & " Activos Fijos que pertenecen a éste.", vbExclamation
         End If
         Call CloseRs(Rs)
         Exit Sub
      End If
   End If
   
   Call CloseRs(Rs)

   
   If MsgBox1("¿Está seguro que desea eliminar el Grupo " & Cb_Grupo & "?" & vbCrLf & vbCrLf & "Atención: Se eliminarán todas las componentes asociadas a este Grupo.", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   Call DeleteSQL(DbMain, "AFGrupos", " WHERE IdGrupo = " & id & " AND IdEmpresa = " & gEmpresa.id)
   
   Call FillGrupo(0)
   
   
End Sub

Private Sub Bt_EditGrupo_Click()
   Dim Frm As FrmGrupo
   Dim Rc As Integer
   Dim id As Long
   Dim Nombre As String
   
   Set Frm = New FrmGrupo
   id = CbItemData(Cb_Grupo)
   Nombre = Cb_Grupo
   Rc = Frm.FEdit(id, Nombre)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Call FillGrupo(id)
   End If

End Sub

Private Sub Bt_NewGrupo_Click()
   Dim Frm As FrmGrupo
   Dim Rc As Integer
   Dim id As Long
   Dim Nombre As String
   
   Set Frm = New FrmGrupo
   Rc = Frm.FNew(id)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Call FillGrupo(id)
   End If
   
End Sub


Private Sub Cb_Grupo_Click()

   Call FillComp(0)
   
End Sub

Private Sub Form_Load()

   Call FillGrupo(0)
    
End Sub

Private Sub FillGrupo(ByVal id As Long)

   Cb_Grupo.Clear

   Call FillCombo(Cb_Grupo, DbMain, "SELECT NombGrupo, IdGrupo FROM AFGrupos WHERE IdEmpresa = " & gEmpresa.id & " ORDER BY NombGrupo", id)
   If id = 0 And Cb_Grupo.ListCount > 0 Then
      Cb_Grupo.ListIndex = 0
   End If

End Sub
Private Sub FillComp(ByVal id As Long)

   Ls_Componentes.Clear

   Call FillCombo(Ls_Componentes, DbMain, "SELECT NombComp, IdComp FROM AFComponentes WHERE IdGrupo = " & CbItemData(Cb_Grupo) & " ORDER BY NombComp", id)

   If id = 0 And Ls_Componentes.ListCount > 0 Then
      Ls_Componentes.ListIndex = 0
   End If

End Sub

Private Sub Bt_DelComp_Click()
   Dim id As Long
   Dim Nombre As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Msg As String
   
   id = CbItemData(Ls_Componentes)
   If id <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Q1 = "SELECT Count(*) FROM ActFijoCompsFicha WHERE IdComp = " & id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then
         If vFld(Rs(0)) = 1 Then
            Msg = "un Activo Fijo que utiliza esta componente. Si la elimina, también será borrada de la definición de este activo fijo."
         Else
            Msg = vFld(Rs(0)) & " Activos Fijos que utilizan esta componente." & vbCrLf & "Si la elimina, también será borrada de la definición de estos activos fijos."
         End If
            
         If MsgBox1("ATENCIÓN: Hay " & Msg & vbCrLf & vbCrLf & "¿Está seguro que desea continuar?", vbYesNoCancel + vbQuestion) <> vbYes Then
            Call CloseRs(Rs)
            Exit Sub
         End If
         
      ElseIf MsgBox1("¿Está seguro que desea eliminar la componente " & Ls_Componentes & "?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
         Exit Sub
      
      End If
      
   ElseIf MsgBox1("¿Está seguro que desea eliminar la componente " & Ls_Componentes & "?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   Call CloseRs(Rs)
            
   
   Call DeleteSQL(DbMain, "AFComponentes", " WHERE IdComp = " & id & " AND IdEmpresa = " & gEmpresa.id)
   Call DeleteSQL(DbMain, "ActFijoCompsFicha", " WHERE IdComp = " & id & " AND IdEmpresa = " & gEmpresa.id)
   
   Call FillComp(0)
   
   
End Sub

Private Sub Bt_EditComp_Click()
   Dim Frm As FrmComponente
   Dim Rc As Integer
   Dim id As Long
   Dim Nombre As String
   
   Set Frm = New FrmComponente
   id = CbItemData(Ls_Componentes)
   Nombre = Ls_Componentes
   Rc = Frm.FEdit(id, Nombre)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Call FillComp(id)
   End If

End Sub

Private Sub Bt_NewComp_Click()
   Dim Frm As FrmComponente
   Dim Rc As Integer
   Dim IdGrupo As Long, IdComp As Long
   Dim Nombre As String
   
   IdGrupo = CbItemData(Cb_Grupo)
   If IdGrupo <= 0 Then
      MsgBox1 "Debe seleccionar un grupo antes de agregar componentes.", vbExclamation
      Exit Sub
   End If
   Set Frm = New FrmComponente
   Rc = Frm.FNew(IdGrupo, IdComp)
   Set Frm = Nothing
   
   If Rc = vbOK Then
      Call FillComp(IdComp)
   End If
   
End Sub

Private Sub Ls_Componentes_DblClick()
   Call Bt_EditComp_Click
End Sub
