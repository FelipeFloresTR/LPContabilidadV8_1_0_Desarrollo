VERSION 5.00
Begin VB.Form FrmAFFicha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activo Fijo - Detalle Financiero"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1575
      Index           =   1
      Left            =   1680
      TabIndex        =   37
      Top             =   2100
      Width           =   5595
      Begin VB.TextBox Tx_FechaCompra 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox Tx_FechaIncorporacion 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox Tx_FechaDisponible 
         Height          =   345
         Left            =   2880
         TabIndex        =   3
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton Bt_FechaIncorporacion 
         Height          =   315
         Left            =   4080
         Picture         =   "FrmAFFicha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   660
         Width           =   230
      End
      Begin VB.CommandButton Bt_FechaDisponible 
         Height          =   315
         Left            =   4080
         Picture         =   "FrmAFFicha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1020
         Width           =   230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha compra:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   41
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Incorporación:"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   39
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha disponible para ser utilizado:"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   38
         Top             =   1080
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Activo Fijo"
      Height          =   1695
      Index           =   0
      Left            =   1680
      TabIndex        =   31
      Top             =   300
      Width           =   5595
      Begin VB.TextBox Tx_Cuenta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   44
         Top             =   1140
         Width           =   4455
      End
      Begin VB.TextBox Tx_Cantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   35
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox Tx_Descrip 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   34
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox Cb_Grupo 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   45
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Index           =   12
         Left            =   3780
         TabIndex        =   36
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Descrip:"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   33
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo:"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Bt_Componentes 
      Caption         =   "Componentes"
      Height          =   735
      Left            =   7560
      Picture         =   "FrmAFFicha.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7560
      TabIndex        =   15
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7560
      TabIndex        =   16
      Top             =   780
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Gastos no reconocidos como parte del precio"
      Height          =   2355
      Index           =   3
      Left            =   1680
      TabIndex        =   25
      Top             =   6420
      Width           =   5595
      Begin VB.TextBox Tx_GastoOtrosConceptos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   13
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox Tx_TotalGastos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Tx_ObrasReubic 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Tx_FormacionPers 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   11
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Tx_IVARecuperable 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Otros conceptos:"
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   43
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   300
         X2              =   4320
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Gastos:"
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
         Index           =   11
         Left            =   300
         TabIndex        =   29
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Obras de reubicación:"
         Height          =   195
         Index           =   10
         Left            =   300
         TabIndex        =   28
         Top             =   1020
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Formación del personal:"
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   27
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "IVA Recuperable:"
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   26
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2595
      Index           =   2
      Left            =   1680
      TabIndex        =   18
      Top             =   3720
      Width           =   5595
      Begin VB.TextBox Tx_AdquiOtrosConceptos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   9
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox Tx_PrecioAdquis 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Tx_ObrasAdapt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Tx_Transporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox Tx_DerechosIntern 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Tx_PrecioFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Otros conceptos."
         Height          =   195
         Index           =   15
         Left            =   300
         TabIndex        =   42
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   300
         X2              =   4440
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio de Adquisición:"
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
         Index           =   7
         Left            =   300
         TabIndex        =   23
         Top             =   2040
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Obras de adaptación:"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   22
         Top             =   1260
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   21
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Derechos de internación:"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   20
         Top             =   660
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Precio según factura:"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   19
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   420
      Picture         =   "FrmAFFicha.frx":0B02
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   17
      Top             =   420
      Width           =   885
   End
End
Attribute VB_Name = "FrmAFFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lIdActFijo As Long
Dim lModif As Boolean

Public Function FEdit(ByVal IdActFijo As Long) As Integer
   
   lIdActFijo = IdActFijo
   
   Me.Show vbModal
   
   FEdit = lRc

End Function

Private Sub Bt_Cancel_Click()
   lRc = vbCancel

   Unload Me
End Sub

Private Sub Bt_Componentes_Click()
   Dim Frm As FrmAFCompsFicha
   
   If Cb_Grupo.ListIndex = -1 Then
      MsgBox1 "Debe seleccionar el Grupo antes de continuar.", vbExclamation
      Exit Sub
   End If
   
   If lModif Then
      If MsgBox1("Se grabarán los datos antes de ingresar a las componentes." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
         Exit Sub
      End If
   
      If Not valida() Then
         Exit Sub
      End If
      
      Call SaveAll
   End If
   
   Set Frm = New FrmAFCompsFicha
   Call Frm.FEdit(lIdActFijo)
   Set Frm = Nothing

End Sub

Private Sub Bt_OK_Click()
   If Not valida() Then
      Exit Sub
   End If
   
   Call SaveAll
   
   lRc = vbOK
   
   Unload Me
End Sub

Private Sub Command1_Click()

End Sub


Private Sub Cb_Grupo_Click()
   lModif = True
End Sub

Private Sub Cb_Grupo_DropDown()
   If Cb_Grupo.Locked Then
      MsgBox1 "No es posible cambiar el Grupo de este Activo Fijo, porque tiene componentes asociadas." & vbCrLf & vbCrLf & "Si elimina todas las componentes del Activo Fijo, podrá cambiar el Grupo.", vbExclamation
   End If
End Sub

Private Sub Form_Load()

   Call SetTxRO(Tx_PrecioFactura, True)

   Call FillGrupo(0)
   Call LoadAll
   
   
End Sub

Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT IdGrupo, MovActivoFijo.Descrip, MovActivoFijo.Cantidad, MovActivoFijo.Fecha as FechaCompra, MovActivoFijo.Neto, MovActivoFijo.IVA, MovActivoFijo.IdCuenta, Cuentas.Descripcion as Cuenta, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, FImported "
   Q1 = Q1 & " FROM (MovActivoFijo LEFT JOIN ActFijoFicha ON ActFijoFicha.IdActfijo =  MovActivoFijo.IdActFijo "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovActivoFijo", "ActFijoFicha") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas ON Cuentas.IdCuenta = MovActivoFijo.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovActivoFijo", "Cuentas")
   Q1 = Q1 & " WHERE MovActivoFijo.IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & gEmpresa.id & " AND MovActivoFijo.Ano = " & gEmpresa.Ano
      
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      Call CbSelItem(Cb_Grupo, vFld(Rs("IdGrupo")))
      Tx_Descrip = vFld(Rs("Descrip"))
      Tx_Cuenta = vFld(Rs("Cuenta"))
      Tx_Cantidad = Format(vFld(Rs("Cantidad")), NUMFMT)
      Call SetTxDate(Tx_FechaCompra, vFld(Rs("FechaCompra")))
      
      Tx_PrecioFactura = Format(vFld(Rs("Neto")), NUMFMT)     'trae el neto del activo fijo
      
      Tx_DerechosIntern = Format(vFld(Rs("DerechosIntern")), NUMFMT)
      Tx_Transporte = Format(vFld(Rs("Transporte")), NUMFMT)
      Tx_ObrasAdapt = Format(vFld(Rs("ObrasAdapt")), NUMFMT)
      Tx_AdquiOtrosConceptos = Format(vFld(Rs("AdquiOtrosConceptos")), NUMFMT)
      Tx_PrecioAdquis = Format(vFld(Rs("PrecioAdquis")), NUMFMT)
      
      Tx_IVARecuperable = Format(vFld(Rs("IVARecuperable")), NUMFMT)
      Tx_FormacionPers = Format(vFld(Rs("FormacionPers")), NUMFMT)
      Tx_ObrasReubic = Format(vFld(Rs("ObrasReubic")), NUMFMT)
      Tx_GastoOtrosConceptos = Format(vFld(Rs("GastoOtrosConceptos")), NUMFMT)
      Tx_TotalGastos = Format(vFld(Rs("TotalGastos")), NUMFMT)
      
      Call SetTxDate(Tx_FechaIncorporacion, vFld(Rs("FechaIncorporacion")))
      Call SetTxDate(Tx_FechaDisponible, vFld(Rs("FechaDisponible")))
      
      Call CalcTot
      
   End If
   
   Call EnableForm0(Me, vFld(Rs("FImported")) = 0 And gEmpresa.FCierre = 0)
   Bt_Componentes.Enabled = True
      
   Call CloseRs(Rs)
   
   Q1 = "SELECT IdComp FROM ActFijoCompsFicha WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY IdComp"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then  'tiene componentes
         Cb_Grupo.Locked = True
      End If
   Else
    Cb_Grupo.Locked = False
   End If
   
   Call CloseRs(Rs)
         
   lModif = False
      
End Sub

Private Sub SaveAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdFicha As Long
   Dim FldArray(2) As AdvTbAddNew_t
   
   Q1 = "SELECT IdFicha "
   Q1 = Q1 & " FROM ActFijoFicha"
   Q1 = Q1 & " WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then   'no existe, lo insertamos
      

      FldArray(0).FldName = "IdActFijo"
      FldArray(0).FldValue = lIdActFijo
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "IdEmpresa"
      FldArray(1).FldValue = gEmpresa.id
      FldArray(1).FldIsNum = True
                  
      FldArray(2).FldName = "Ano"
      FldArray(2).FldValue = gEmpresa.Ano
      FldArray(2).FldIsNum = True
            
      IdFicha = AdvTbAddNewMult(DbMain, "ActFijoFicha", "IdFicha", FldArray)
   
'      Set Rs = DbMain.OpenRecordset("ActFijoFicha")
'      Rs.AddNew
'
'      IdFicha = Rs("IdFicha")
'      Rs.Fields("IdActFijo") = lIdActFijo
'
'      Rs.Update
'      Rs.Close
'      Set Rs = Nothing

'      Q1 = "UPDATE ActFijoFicha SET "
'      Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
'      Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'      Q1 = Q1 & " WHERE IdFicha = " & IdFicha
'
'      Call ExecSQL(DbMain, Q1)
      
   Else
      IdFicha = vFld(Rs("IdFicha"))
      
   End If
   
   Call CloseRs(Rs)
      
   Q1 = "UPDATE ActFijoFicha SET "
   
   Q1 = Q1 & "  IdGrupo = " & CbItemData(Cb_Grupo)
   Q1 = Q1 & ", FechaIncorporacion = " & GetTxDate(Tx_FechaIncorporacion)
   Q1 = Q1 & ", FechaDisponible = " & GetTxDate(Tx_FechaDisponible)
   
   Q1 = Q1 & ", PrecioFactura = " & vFmt(Tx_PrecioFactura)
   Q1 = Q1 & ", DerechosIntern = " & vFmt(Tx_DerechosIntern)
   Q1 = Q1 & ", Transporte = " & vFmt(Tx_Transporte)
   Q1 = Q1 & ", ObrasAdapt = " & vFmt(Tx_ObrasAdapt)
   Q1 = Q1 & ", AdquiOtrosConceptos = " & vFmt(Tx_AdquiOtrosConceptos)
   Q1 = Q1 & ", PrecioAdquis = " & vFmt(Tx_PrecioAdquis)
   
   Q1 = Q1 & ", IVARecuperable = " & vFmt(Tx_IVARecuperable)
   Q1 = Q1 & ", FormacionPers = " & vFmt(Tx_FormacionPers)
   Q1 = Q1 & ", ObrasReubic = " & vFmt(Tx_ObrasReubic)
   Q1 = Q1 & ", GastoOtrosConceptos = " & vFmt(Tx_GastoOtrosConceptos)
   Q1 = Q1 & ", TotalGastos = " & vFmt(Tx_TotalGastos)
      
   Q1 = Q1 & " WHERE IdFicha = " & IdFicha
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Call ExecSQL(DbMain, Q1)
   
   lModif = False


End Sub

Private Function valida()
   valida = False
   
   If Cb_Grupo.ListIndex = -1 Then
      MsgBox1 "Debe seleccionar el Grupo antes de continuar.", vbExclamation
      Exit Function
   End If

   
   If vFmt(Tx_PrecioFactura) = 0 Then
      MsgBox1 "Precio Factura inválido.", vbExclamation
      Exit Function
   End If
   
   If GetTxDate(Tx_FechaIncorporacion) > 0 And GetTxDate(Tx_FechaIncorporacion) < GetTxDate(Tx_FechaCompra) Then
      MsgBox1 "Fecha de incorporación anterior a fecha de compra.", vbExclamation
      Exit Function
   End If
   
   If GetTxDate(Tx_FechaDisponible) > 0 And GetTxDate(Tx_FechaDisponible) < GetTxDate(Tx_FechaCompra) Then
      MsgBox1 "Fecha disponible anterior a fecha de compra.", vbExclamation
      Exit Function
   End If
   
   valida = True

End Function



Private Sub Tx_AdquiOtrosConceptos_Change()
   lModif = True

End Sub

Private Sub Tx_AdquiOtrosConceptos_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_AdquiOtrosConceptos_LostFocus()
   Tx_AdquiOtrosConceptos = Format(vFmt(Tx_AdquiOtrosConceptos), NUMFMT)
   Call CalcTot

End Sub
Private Sub Tx_GastoOtrosConceptos_Change()
   lModif = True

End Sub

Private Sub Tx_GastoOtrosConceptos_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_GastoOtrosConceptos_LostFocus()
   Tx_GastoOtrosConceptos = Format(vFmt(Tx_GastoOtrosConceptos), NUMFMT)
   Call CalcTot

End Sub

Private Sub Tx_DerechosIntern_Change()
   lModif = True

End Sub

Private Sub Tx_DerechosIntern_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_DerechosIntern_LostFocus()
   Tx_DerechosIntern = Format(vFmt(Tx_DerechosIntern), NUMFMT)
   Call CalcTot

End Sub

Private Sub Tx_FechaDisponible_Change()
   lModif = True

End Sub

Private Sub Tx_FechaIncorporacion_Change()
   lModif = True

End Sub

Private Sub Tx_FormacionPers_Change()
   lModif = True

End Sub

Private Sub Tx_FormacionPers_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_FormacionPers_LostFocus()
   Tx_FormacionPers = Format(vFmt(Tx_FormacionPers), NUMFMT)
   
   Call CalcTot

End Sub

Private Sub Tx_IVARecuperable_Change()
   lModif = True

End Sub

Private Sub Tx_IVARecuperable_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_IVARecuperable_LostFocus()
   Tx_IVARecuperable = Format(vFmt(Tx_IVARecuperable), NUMFMT)
   
   Call CalcTot

End Sub

Private Sub Tx_ObrasAdapt_Change()
   lModif = True

End Sub

Private Sub Tx_ObrasAdapt_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_ObrasAdapt_LostFocus()
   Tx_ObrasAdapt = Format(vFmt(Tx_ObrasAdapt), NUMFMT)
   
   Call CalcTot

End Sub

Private Sub Tx_ObrasReubic_Change()
   lModif = True

End Sub

Private Sub Tx_ObrasReubic_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_ObrasReubic_LostFocus()
   Tx_ObrasReubic = Format(vFmt(Tx_ObrasReubic), NUMFMT)

   Call CalcTot

End Sub

Private Sub Tx_PrecioFactura_Change()
   lModif = True
End Sub

Private Sub Tx_PrecioFactura_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_PrecioFactura_LostFocus()
   Tx_PrecioFactura = Format(vFmt(Tx_PrecioFactura), NUMFMT)
   
   Call CalcTot
End Sub

Private Sub Tx_Transporte_Change()
   lModif = True

End Sub

Private Sub Tx_Transporte_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
Private Sub Tx_FechaIncorporacion_GotFocus()
   Call DtGotFocus(Tx_FechaIncorporacion)
End Sub


Private Sub Tx_FechaIncorporacion_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_FechaIncorporacion) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_FechaIncorporacion)
   
   Call DtLostFocus(Tx_FechaIncorporacion)
      
End Sub

Private Sub Tx_FechaIncorporacion_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub
Private Sub Bt_FechaIncorporacion_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaIncorporacion)
      
   Set Frm = Nothing
   
   lModif = True

End Sub

Private Sub Tx_FechaDisponible_GotFocus()
   Call DtGotFocus(Tx_FechaDisponible)
End Sub


Private Sub Tx_FechaDisponible_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_FechaDisponible) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_FechaDisponible)
   
   Call DtLostFocus(Tx_FechaDisponible)
      
End Sub

Private Sub Tx_FechaDisponible_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub
Private Sub Bt_FechaDisponible_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaDisponible)
      
   Set Frm = Nothing
   
   lModif = True

End Sub

Private Sub CalcTot()
   Dim Tot As Double
      
   Tot = vFmt(Tx_PrecioFactura) + vFmt(Tx_DerechosIntern) + vFmt(Tx_Transporte) + vFmt(Tx_ObrasAdapt) + vFmt(Tx_AdquiOtrosConceptos)
   
   Tx_PrecioAdquis = Format(Tot, NUMFMT)
   
   Tot = vFmt(Tx_IVARecuperable) + vFmt(Tx_FormacionPers) + vFmt(Tx_ObrasReubic) + vFmt(Tx_GastoOtrosConceptos)
   
   Tx_TotalGastos = Format(Tot, NUMFMT)
   
End Sub

Private Sub Tx_Transporte_LostFocus()
   Tx_Transporte = Format(vFmt(Tx_Transporte), NUMFMT)
   
   Call CalcTot

End Sub
Private Sub FillGrupo(ByVal id As Long)

   Cb_Grupo.Clear

   Call FillCombo(Cb_Grupo, DbMain, "SELECT NombGrupo, IdGrupo FROM AFGrupos WHERE IdEmpresa = " & gEmpresa.id & " ORDER BY NombGrupo", id)
   If id = 0 And Cb_Grupo.ListCount > 0 Then
      Cb_Grupo.ListIndex = 0
   End If

End Sub

