VERSION 5.00
Begin VB.Form FrmContEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Empresa"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Procesos Anuales"
      Height          =   4995
      Index           =   0
      Left            =   4680
      TabIndex        =   4
      Top             =   360
      Width           =   3375
      Begin VB.Frame Frame2 
         Caption         =   "Activo Fijo"
         Height          =   1455
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   420
         Width           =   3015
         Begin VB.CheckBox Ch_Depreciacion 
            Alignment       =   1  'Right Justify
            Caption         =   "Depreciación"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   300
            Width           =   2235
         End
         Begin VB.CheckBox Ch_CM 
            Alignment       =   1  'Right Justify
            Caption         =   "Corrección Monetaria"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   660
            Width           =   2235
         End
         Begin VB.CheckBox Ch_33Bis 
            Alignment       =   1  'Right Justify
            Caption         =   "33 Bis LIR"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   1020
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Corrección Monetaria"
         Height          =   1155
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   2040
         Width           =   3015
         Begin VB.CheckBox Ch_Pasivos 
            Alignment       =   1  'Right Justify
            Caption         =   "Pasivos"
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   660
            Width           =   2235
         End
         Begin VB.CheckBox Ch_Activos 
            Alignment       =   1  'Right Justify
            Caption         =   "Activos"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   300
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   3360
         Width           =   3015
         Begin VB.CheckBox Ch_F22Renta 
            Alignment       =   1  'Right Justify
            Caption         =   "F22 Renta"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1020
            Width           =   2235
         End
         Begin VB.CheckBox Ch_CPTMunicip 
            Alignment       =   1  'Right Justify
            Caption         =   "CPT Municipalidad"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   660
            Width           =   2235
         End
         Begin VB.CheckBox Ch_BalDefinitivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Balance Definitivo"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   300
            Width           =   2235
         End
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8400
      TabIndex        =   2
      Top             =   780
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8400
      TabIndex        =   1
      Top             =   420
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodos Contables Cerrados"
      Height          =   5055
      Left            =   1140
      TabIndex        =   3
      Top             =   360
      Width           =   3255
      Begin VB.ListBox Ls_Meses 
         Height          =   2985
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   540
         Width           =   2835
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   1
      Left            =   300
      Picture         =   "FrmContEmpresa.frx":0000
      Top             =   480
      Width           =   690
   End
End
Attribute VB_Name = "FrmContEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   Call SaveAll
   Unload Me
End Sub

Private Sub Form_Load()
   Call LoadAll
   
   If Not ChkPriv(PRV_ADM_EMPRESA) Then
      Bt_Ok.Enabled = False
   End If
      
End Sub

Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim Idx As Long
   
   'insertamos los meses sin check
   For i = 1 To 12
      Call AddItem(Ls_Meses, gNomMes(i), i)
   Next i
   
   'ahora consultamos los datos
   Q1 = "SELECT "
   For i = 1 To 12
      Q1 = Q1 & "Mes" & i & ","
   Next i

   Q1 = Q1 & " AF_Depreciacion, AF_CM, AF_33BisLir, CM_Activos, CM_Pasivos, BalDefinitivo, CPT_Municip, F22Renta"
   Q1 = Q1 & " FROM ControlEmpresa"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      For i = 1 To 12
      
         If vFld(Rs("Mes" & i)) <> 0 Then
            Idx = FindItem(Ls_Meses, i)
            If Idx >= 0 Then
               Ls_Meses.Selected(Idx) = True
            End If
         End If
         
      Next i
      
      Ch_Depreciacion = IIf(vFld(Rs("AF_Depreciacion")) <> 0, 1, 0)
      Ch_CM = IIf(vFld(Rs("AF_CM")) <> 0, 1, 0)
      Ch_33Bis = IIf(vFld(Rs("AF_33BisLir")) <> 0, 1, 0)
      Ch_Activos = IIf(vFld(Rs("CM_Activos")) <> 0, 1, 0)
      Ch_Pasivos = IIf(vFld(Rs("CM_Pasivos")) <> 0, 1, 0)
      Ch_BalDefinitivo = IIf(vFld(Rs("BalDefinitivo")) <> 0, 1, 0)
      Ch_CPTMunicip = IIf(vFld(Rs("CPT_Municip")) <> 0, 1, 0)
      Ch_F22Renta = IIf(vFld(Rs("F22Renta")) <> 0, 1, 0)
      
   End If
   
   Call CloseRs(Rs)

End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer

   'veamos si está el registro
   Q1 = "SELECT * FROM ControlEmpresa WHERE IdEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then
      Q1 = "INSERT INTO ControlEmpresa (IdEmpresa, Ano, RUT, RazonSocial) "
      Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & gEmpresa.Rut & ",'" & ParaSQL(gEmpresa.RazonSocial) & "')"
      Call ExecSQL(DbMain, Q1)
   End If
   
   Call CloseRs(Rs)
   
   'ahora actualizamos
   Q1 = "UPDATE ControlEmpresa SET "
   
   For i = 0 To Ls_Meses.ListCount - 1
   
      Q1 = Q1 & "Mes" & Ls_Meses.ItemData(i) & "=" & IIf(Ls_Meses.Selected(i), 1, 0) & ","
      
   Next i
      
   Q1 = Q1 & "  AF_Depreciacion = " & IIf(Ch_Depreciacion <> 0, 1, 0)
   Q1 = Q1 & ", AF_CM = " & IIf(Ch_CM <> 0, 1, 0)
   Q1 = Q1 & ", AF_33BisLir = " & IIf(Ch_33Bis <> 0, 1, 0)
   Q1 = Q1 & ", CM_Activos = " & IIf(Ch_Activos <> 0, 1, 0)
   Q1 = Q1 & ", CM_Pasivos = " & IIf(Ch_Pasivos <> 0, 1, 0)
   Q1 = Q1 & ", BalDefinitivo = " & IIf(Ch_BalDefinitivo <> 0, 1, 0)
   Q1 = Q1 & ", CPT_Municip = " & IIf(Ch_CPTMunicip <> 0, 1, 0)
   Q1 = Q1 & ", F22Renta = " & IIf(Ch_F22Renta <> 0, 1, 0)
   
   Q1 = Q1 & " WHERE IdEmpresa= " & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
End Sub
