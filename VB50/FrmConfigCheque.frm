VERSION 5.00
Begin VB.Form FrmConfigCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Impresión de Cheques"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "FrmConfigCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   180
      Picture         =   "FrmConfigCheque.frx":000C
      ScaleHeight     =   720
      ScaleWidth      =   780
      TabIndex        =   30
      Top             =   480
      Width           =   840
   End
   Begin VB.ComboBox Cb_Tipo 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cheque nominativo"
      Height          =   1155
      Left            =   1320
      TabIndex        =   19
      Top             =   6420
      Width           =   5115
      Begin VB.CheckBox Ch_BorrarAlPortador 
         Caption         =   "Borrar al Portador"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1995
      End
      Begin VB.CheckBox Ch_BorrarAlaOrden 
         Caption         =   "Borrar a la Orden De "
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.CommandButton Bt_Test 
      Caption         =   "Imprimir marca"
      Height          =   675
      Left            =   6960
      Picture         =   "FrmConfigCheque.frx":0575
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime marca en el borde superior izquierdo del cheque"
      Top             =   3300
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6960
      TabIndex        =   15
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6960
      TabIndex        =   14
      Top             =   240
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Posición esquina superior izquierda cheque"
      Height          =   5415
      Left            =   1260
      TabIndex        =   16
      Top             =   840
      Width           =   5175
      Begin VB.TextBox Tx_AnoMove 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Distancia a desplazar hacia la derecha (si es negativo, hacia Izquierda) en twips"
         Top             =   3600
         Width           =   1275
      End
      Begin VB.CheckBox Ch_Omitir2DigAno 
         Caption         =   "Omitir primeros dos dígitos en el año"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   3300
         Width           =   3855
      End
      Begin VB.TextBox Tx_OrdenDeMove 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   10
         ToolTipText     =   "Distancia a desplazar hacia la derecha (si es negativo, hacia Izquierda) en twips"
         Top             =   4500
         Width           =   1275
      End
      Begin VB.TextBox Tx_FechaMove 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   6
         ToolTipText     =   "Distancia a desplazar hacia la derecha (si es negativo, hacia Izquierda) en twips"
         Top             =   2820
         Width           =   1275
      End
      Begin VB.TextBox Tx_ValDigMove 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   4
         ToolTipText     =   "Distancia a desplazar hacia la derecha (si es negativo, hacia Izquierda) en twips"
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox Tx_OrdenDe 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   9
         ToolTipText     =   "Distancia a desplazar hacia Abajo (si es negativo, hacia Arriba) en twips"
         Top             =   4140
         Width           =   1275
      End
      Begin VB.TextBox Tx_Fecha 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   5
         ToolTipText     =   "Distancia a desplazar hacia Abajo (si es negativo, hacia Arriba) en twips"
         Top             =   2460
         Width           =   1275
      End
      Begin VB.TextBox Tx_ValDig 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   3
         ToolTipText     =   "Distancia a desplazar hacia Abajo (si es negativo, hacia Arriba) en twips"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox Tx_Izq 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   2
         ToolTipText     =   "Distancia en twips desde el borde izquierdo de la hoja"
         Top             =   900
         Width           =   1275
      End
      Begin VB.TextBox Tx_Sup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Distancia en Twips desde el borde inferior de la hoja"
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mover a izq./der. año:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   38
         Top             =   3660
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Index           =   2
         Left            =   4500
         TabIndex        =   37
         Top             =   3660
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   36
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Mover izq./der. Orden De:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   4560
         Width           =   1860
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   34
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mover a izq./der. Fecha:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   33
         Top             =   2880
         Width           =   1755
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Left            =   4500
         TabIndex        =   32
         Top             =   1860
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Mover a izq./der. Valor:"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   1860
         Width           =   1665
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Bajar Orden De en;"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   4200
         Width           =   1365
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   28
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Bajar Fecha en:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   26
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bajar Valor en dígitos:"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   1500
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Left            =   4500
         TabIndex        =   24
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nota: 1 cm equivale a 567 twips"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   4980
         Width           =   2280
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Left            =   4500
         TabIndex        =   22
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "twips"
         Height          =   195
         Left            =   4500
         TabIndex        =   21
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Borde izquierdo cheque:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Borde superior cheque:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Width           =   1650
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo papel:"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   300
      Width           =   795
   End
End
Attribute VB_Name = "FrmConfigCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_OK_Click()

   Call SetIniString(gIniFile, "Cheques", "TipoPapel", CbItemData(Cb_Tipo))
   gPrtCheques.TipoPapel = CbItemData(Cb_Tipo)
      
   If CbItemData(Cb_Tipo) = CHEQUE_CARTA Then           'hoja carta
   
      Call SetIniString(gIniFile, "Cheques", "Altura", vFmt(Tx_Sup))
      Call SetIniString(gIniFile, "Cheques", "BordeIzq", vFmt(Tx_Izq))
      Call SetIniString(gIniFile, "Cheques", "BajarValDig", vFmt(Tx_ValDig))
      Call SetIniString(gIniFile, "Cheques", "MoverValDig", vFmt(Tx_ValDigMove))
      Call SetIniString(gIniFile, "Cheques", "BajarFecha", vFmt(Tx_Fecha))
      Call SetIniString(gIniFile, "Cheques", "MoverFecha", vFmt(Tx_FechaMove))
      Call SetIniString(gIniFile, "Cheques", "Omitir2DigAno", IIf(Ch_Omitir2DigAno <> 0, 1, 0))
      Call SetIniString(gIniFile, "Cheques", "MoverAno", vFmt(Tx_AnoMove))
      Call SetIniString(gIniFile, "Cheques", "BajarOrdenDe", vFmt(Tx_OrdenDe))
      Call SetIniString(gIniFile, "Cheques", "MoverOrdenDe", vFmt(Tx_OrdenDeMove))
      Call SetIniString(gIniFile, "Cheques", "BorrarOrden", Ch_BorrarAlaOrden)
      Call SetIniString(gIniFile, "Cheques", "BorrarPortador", Ch_BorrarAlPortador)
   
      gPrtCheques.AlturaCheque = vFmt(Tx_Sup)
      gPrtCheques.BordeIzqCheque = vFmt(Tx_Izq)
      gPrtCheques.BajarValDig = vFmt(Tx_ValDig)
      gPrtCheques.MoverValDig = vFmt(Tx_ValDigMove)
      gPrtCheques.BajarFecha = vFmt(Tx_Fecha)
      gPrtCheques.MoverFecha = vFmt(Tx_FechaMove)
      gPrtCheques.Omitir2DigAno = IIf(Ch_Omitir2DigAno <> 0, 1, 0)
      gPrtCheques.MoverAno = vFmt(Tx_AnoMove)
      gPrtCheques.BajarOrdenDe = vFmt(Tx_OrdenDe)
      gPrtCheques.MoverOrdenDe = vFmt(Tx_OrdenDeMove)
      
      
      gPrtCheques.BorrarALaOrden = Ch_BorrarAlaOrden <> 0
      gPrtCheques.BorrarAlPortador = Ch_BorrarAlPortador <> 0
   
   Else
   
      Call SetIniString(gIniFile, "Cheques", "PCont-BordeSup", vFmt(Tx_Sup))
      Call SetIniString(gIniFile, "Cheques", "PCont-BordeIzq", vFmt(Tx_Izq))
      Call SetIniString(gIniFile, "Cheques", "PCont-BajarValDig", vFmt(Tx_ValDig))
      Call SetIniString(gIniFile, "Cheques", "PCont-MoverValDig", vFmt(Tx_ValDigMove))
      Call SetIniString(gIniFile, "Cheques", "PCont-BajarFecha", vFmt(Tx_Fecha))
      Call SetIniString(gIniFile, "Cheques", "PCont-MoverFecha", vFmt(Tx_FechaMove))
      Call SetIniString(gIniFile, "Cheques", "PCont-Omitir2DigAno", IIf(Ch_Omitir2DigAno <> 0, 1, 0))
      Call SetIniString(gIniFile, "Cheques", "PCont-MoverAno", vFmt(Tx_AnoMove))
      Call SetIniString(gIniFile, "Cheques", "PCont-BajarOrdenDe", vFmt(Tx_OrdenDe))
      Call SetIniString(gIniFile, "Cheques", "PCont-MoverOrdenDe", vFmt(Tx_OrdenDeMove))
      Call SetIniString(gIniFile, "Cheques", "PCont-BorrarOrden", Ch_BorrarAlaOrden)
      Call SetIniString(gIniFile, "Cheques", "PCont-BorrarPortador", Ch_BorrarAlPortador)
      
      gPrtCheques.BordeSuperiorPCont = vFmt(Tx_Sup)
      gPrtCheques.BordeIzqChequePCont = vFmt(Tx_Izq)
      gPrtCheques.BajarValDigPCont = vFmt(Tx_ValDig)
      gPrtCheques.MoverValDigPCont = vFmt(Tx_ValDigMove)
      gPrtCheques.BajarFechaPCont = vFmt(Tx_Fecha)
      gPrtCheques.MoverFechaPCont = vFmt(Tx_FechaMove)
      gPrtCheques.Omitir2DigAnoPCont = IIf(Ch_Omitir2DigAno <> 0, 1, 0)
      gPrtCheques.MoverAnoPCont = vFmt(Tx_AnoMove)
      gPrtCheques.BajarOrdenDePCont = vFmt(Tx_OrdenDe)
      gPrtCheques.MoverOrdenDePCont = vFmt(Tx_OrdenDeMove)
      
      gPrtCheques.BorrarALaOrdenPCont = Ch_BorrarAlaOrden <> 0
      gPrtCheques.BorrarAlPortadorPCont = Ch_BorrarAlPortador <> 0
   
   End If

   
   Unload Me
End Sub

Private Sub bt_Test_Click()
   If Not PrepararPrt(FrmMain.Cm_PrtDlg) Then
      Exit Sub
   End If
   
   Call gPrtCheques.PrtMarca(Printer)
End Sub

Private Sub Cb_Tipo_Click()

   If CbItemData(Cb_Tipo) = CHEQUE_CARTA Then
      Tx_Sup.ToolTipText = "Distancia en Twips desde el borde inferior de la hoja"
   Else
      Tx_Sup.ToolTipText = "Distancia en Twips desde el borde superior del cheque"
   End If
   
   Call LoadAll
   
End Sub

Private Sub Ch_BorrarAlaOrden_Click()
   Cb_Tipo.Locked = True

End Sub

Private Sub Ch_BorrarAlPortador_Click()
   Cb_Tipo.Locked = True

End Sub

Private Sub Form_Load()
   
   Call CbAddItem(Cb_Tipo, "Hoja carta", CHEQUE_CARTA)
   Call CbAddItem(Cb_Tipo, "Papel continuo", CHEQUE_CONTINUO)
'   Cb_Tipo.ListIndex = 0
   
   Call CbSelItem(Cb_Tipo, gPrtCheques.TipoPapel)
   
   Call LoadAll
      
End Sub

Private Sub LoadAll()
   Dim Altura As Long
   Dim BordeIzq As Long
   Dim BordeSup As Long
   Dim BajarValDig As Long
   Dim MoverValDig As Long
   Dim BajarFecha As Long
   Dim MoverFecha As Long
   Dim Omitir2DigAno As Long
   Dim MoverAno As Long
   Dim BajarOrdenDe As Long
   Dim MoverOrdenDe As Long
   Dim BorrarALaOrden As Long
   Dim BorrarAlPortador As Long
   Dim Tipo As Long
   
   Tipo = CbItemData(Cb_Tipo)
   
   If Tipo = CHEQUE_CARTA Then        'Hoja carta
        
      Tx_Sup = gPrtCheques.AlturaCheque
      Tx_Izq = gPrtCheques.BordeIzqCheque
      
      Altura = Val(GetIniString(gIniFile, "Cheques", "Altura", ""))
      BordeIzq = Val(GetIniString(gIniFile, "Cheques", "BordeIzq", ""))
      BajarValDig = Val(GetIniString(gIniFile, "Cheques", "BajarValDig", ""))
      MoverValDig = Val(GetIniString(gIniFile, "Cheques", "MoverValDig", ""))
      BajarFecha = Val(GetIniString(gIniFile, "Cheques", "BajarFecha", ""))
      MoverFecha = Val(GetIniString(gIniFile, "Cheques", "MoverFecha", ""))
      MoverAno = Val(GetIniString(gIniFile, "Cheques", "MoverAno", ""))
      Omitir2DigAno = Val(GetIniString(gIniFile, "Cheques", "Omitir2DigAno", 0))
      BajarOrdenDe = Val(GetIniString(gIniFile, "Cheques", "BajarOrdenDe", ""))
      MoverOrdenDe = Val(GetIniString(gIniFile, "Cheques", "MoverOrdenDe", ""))
      
      
      BorrarALaOrden = Val(GetIniString(gIniFile, "Cheques", "BorrarOrden", ""))
      BorrarAlPortador = Val(GetIniString(gIniFile, "Cheques", "BorrarPortador", ""))
      
   Else           'papel continuo

      Tx_Sup = gPrtCheques.BordeSuperiorPCont
      Tx_Izq = gPrtCheques.BordeIzqChequePCont
      
      BordeSup = Val(GetIniString(gIniFile, "Cheques", "PCont-BordeSup", ""))
      BordeIzq = Val(GetIniString(gIniFile, "Cheques", "PCont-BordeIzq", ""))
      BajarValDig = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarValDig", ""))
      MoverValDig = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverValDig", ""))
      BajarFecha = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarFecha", ""))
      MoverFecha = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverFecha", ""))
      MoverAno = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverAno", ""))
      Omitir2DigAno = Val(GetIniString(gIniFile, "Cheques", "PCont-Omitir2DigAno", 0))
      BajarOrdenDe = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarOrdenDe", ""))
      MoverOrdenDe = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverOrdenDe", ""))
      
      BorrarALaOrden = Val(GetIniString(gIniFile, "Cheques", "PCont-BorrarOrden", ""))
      BorrarAlPortador = Val(GetIniString(gIniFile, "Cheques", "PCont-BorrarPortador", ""))
      
   End If
   
   If Altura <> 0 Then
      Tx_Sup = Altura
   End If
   If BordeIzq <> 0 Then
      Tx_Izq = BordeIzq
   End If
   If BajarValDig <> 0 Then
      Tx_ValDig = BajarValDig
   End If
   If MoverValDig <> 0 Then
      Tx_ValDigMove = MoverValDig
   End If
   If BajarFecha <> 0 Then
      Tx_Fecha = BajarFecha
   End If
   If MoverFecha <> 0 Then
      Tx_FechaMove = MoverFecha
   End If
   If MoverAno <> 0 Then
      Tx_AnoMove = MoverAno
   End If
   If Omitir2DigAno <> 0 Then
      Ch_Omitir2DigAno = 1
   End If
   If BajarOrdenDe <> 0 Then
      Tx_OrdenDe = BajarOrdenDe
   End If
   If MoverOrdenDe <> 0 Then
      Tx_OrdenDeMove = MoverOrdenDe
   End If
   
   Ch_BorrarAlaOrden = Abs(BorrarALaOrden <> 0)
   Ch_BorrarAlPortador = Abs(BorrarAlPortador <> 0)

  
   Cb_Tipo.Locked = False
   
End Sub

Private Sub Tx_Izq_Change()
   Cb_Tipo.Locked = True

End Sub

Private Sub Tx_Izq_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)

End Sub

Private Sub Tx_Sup_Change()
   Cb_Tipo.Locked = True
End Sub

Private Sub Tx_Sup_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_ValDig_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_ValDig_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_ValDigMove_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_ValDigMove_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_Fecha_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_Fecha_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_FechaMove_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_FechaMove_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Sub Tx_AnoMove_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_AnoMove_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Sub Tx_OrdenDe_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_OrdenDe_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Sub Tx_OrdenDeMove_Change()
   Cb_Tipo.Locked = True

End Sub
Private Sub Tx_OrdenDeMove_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Sub Ch_Omitir2DigAno_Click()
   Cb_Tipo.Locked = True

End Sub

