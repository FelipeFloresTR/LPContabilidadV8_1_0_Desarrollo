VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmApertura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Comprobante de Apertura Año"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "FrmApertura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10020
      TabIndex        =   5
      Top             =   900
      Width           =   1395
   End
   Begin VB.CommandButton Bt_AperturaAno 
      Caption         =   "Generar"
      Height          =   315
      Left            =   10020
      TabIndex        =   4
      Top             =   540
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1500
      TabIndex        =   10
      Top             =   3360
      Width           =   8175
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprobante de Apertura"
      Height          =   2415
      Left            =   1500
      TabIndex        =   6
      Top             =   480
      Width           =   8175
      Begin VB.TextBox Tx_RemIVAUTMAnoAnt 
         Height          =   315
         Left            =   2820
         TabIndex        =   3
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox Tx_CtaCredIVA 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   4695
      End
      Begin VB.CommandButton Bt_CtaCredIVA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7500
         Picture         =   "FrmApertura.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Plan de Cuentas"
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox Tx_CompAper 
         Height          =   315
         Left            =   2820
         TabIndex        =   0
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Cuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7500
         Picture         =   "FrmApertura.frx":03E0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Plan de Cuentas"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox Tx_Cuenta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   780
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remanente IVA año anterior    UTM"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2550
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta de Crédito IVA:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Lb_Reemp 
         Caption         =   "Se reemplazará Comprobante de Apertura ya existente."
         Height          =   435
         Left            =   4260
         TabIndex        =   12
         Top             =   270
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de comprobante de apertura:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta de resultado:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1635
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   360
      Picture         =   "FrmApertura.frx":07B4
      Top             =   600
      Width           =   780
   End
End
Attribute VB_Name = "FrmApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COD_CTARESULT = "2031101"
Const DESC_CTARESULT = "Utilidad Neta Retenida"
Const COD_CTACREDIVA = "1010999"
Const DESC_CTACREDIVA = "Otros Impuestos por Recuperar"


Dim lAno As Long
Dim lIdEmpresa As Long
Dim lidCuentaResul As Long
Dim lidCuentaCredIVA As Long
Dim lIdCompAper As Long
Dim lIdCompAperTrib As Long
Dim lNumCompAper As Long
Dim lParamCtaResult As Boolean
Dim lParamCtaCredIVA As Boolean

Dim lRc As Integer

Public Function FSelect(ByVal IdEmpresa As Long, ByVal Ano As Integer, NumCompAper As Long, IdCompAper As Long, IdCuentaResul As Long, IdCompAperTrib As Long) As Integer
   lIdEmpresa = IdEmpresa
   lAno = Ano
   
   Me.Show vbModal
   
   
   If lRc = vbOK Then
      NumCompAper = lNumCompAper
      IdCompAper = lIdCompAper
      IdCuentaResul = lidCuentaResul
      IdCompAperTrib = lIdCompAperTrib
   End If
   
   FSelect = lRc
   
 
End Function

'14690904
Public Function FSelectDuplicados(ByVal IdEmpresa As Long, ByVal Ano As Integer, NumCompAper As Long, IdCompAper As Long, IdCuentaResul As Long, IdCompAperTrib As Long) As Integer
   lIdEmpresa = IdEmpresa
   lAno = Ano
   
   'Me.Show vbModal
   
  Call Bt_AperturaAno_Click
   
   If lRc = vbOK Then
      NumCompAper = lNumCompAper
      IdCompAper = lIdCompAper
      IdCuentaResul = lidCuentaResul
      IdCompAperTrib = lIdCompAperTrib
   End If
   
   FSelectDuplicados = lRc
   
 
End Function
'14690904

Private Sub Bt_Add_Click()
   Dim Frm As FrmPlanCuentas
   
   MousePointer = vbHourglass
   Set Frm = New FrmPlanCuentas
   Call Frm.FEdit
   Set Frm = Nothing
   MousePointer = vbDefault

End Sub

Private Sub Bt_AperturaAno_Click()
   Dim Msg As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim F1 As Long
   Dim RutMdb As String
   Dim idCompApertura As Long
   Dim TblNew As Boolean
   Dim i As Integer
   
   If valida() = False Then
      Exit Sub
   End If
         
   lNumCompAper = Tx_CompAper
   
   Call SaveCuentas
   
   i = ProgressBar.Value
   If i < 100 Then
      ProgressBar.Value = 100
   End If
             
   lRc = vbOK
   Unload Me
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me

End Sub

Private Sub Bt_CtaCredIVA_Click()
   Dim Frm As FrmPlanCuentas
   Dim DescCta As String
   Dim Codigo As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas
   If Frm.FSelEdit(lidCuentaCredIVA, Codigo, DescCta, Nombre) = vbOK Then
      Tx_CtaCredIVA = DescCta
   End If
   Set Frm = Nothing

End Sub

Private Sub Bt_Cuentas_Click()
   Dim Frm As FrmPlanCuentas
   Dim DescCta As String
   Dim Codigo As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas
   If Frm.FSelEdit(lidCuentaResul, Codigo, DescCta, Nombre) = vbOK Then
      Tx_Cuenta = DescCta
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim Rs As Recordset
   
   'Call EnableForm(Me, gEmpresa.FCierre = 0)
     
   Me.Caption = Me.Caption & " " & lAno

   Call SetTxRO(Tx_Cuenta, True) 'La vuelvo a bloquear, porque se desbloqueo con el EnableForm
   
   Call LoadAll
   
   Call SetupPriv
   
End Sub
Private Function valida() As Boolean
   
   valida = False
   If vFmt(Tx_CompAper) = 0 Then
      MsgBox1 "Debe ingresar número de comprobante de apertura para el año " & lAno, vbExclamation
      Tx_CompAper.SetFocus
      Exit Function
   End If
      
   If Trim(Tx_Cuenta) = "" Then
      MsgBox1 "Debe seleccionar una cuenta de patrimonio para el comprobante de apertura año " & lAno, vbExclamation
      Bt_Cuentas.SetFocus
      Exit Function
      
   End If
   
   If Trim(Tx_CtaCredIVA) = "" Then
      MsgBox1 "Debe seleccionar una cuenta de arrastre de Crédito IVA desde el año anterior.", vbExclamation
      Bt_Cuentas.SetFocus
      Exit Function
      
   End If
   
   valida = True
      

End Function
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim NumLastComp(N_TIPOCOMP) As Long
   Dim IdCompAper As Long, IdCompAperTrib As Long
   Dim NumCompAper As Long, NumCompAperTrib As Long
   Dim HayAnoAnterior As Boolean
   Dim Wh As String
   Dim RemIVAUTM As Double
   Dim RemIVAUTMAnoAnt As Double
     
   
   'obtengo Id Comp. Apertura generado automáticamente, si ya hay uno, para reemplazarlo
   
   Q1 = "SELECT IdCompAper, NCompAper, IdCompAperTrib, NCompAperTrib, RemIVAUTMAnoAnt "
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE idEmpresa=" & lIdEmpresa & " AND Ano=" & lAno
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      IdCompAper = vFld(Rs("IdCompAper"))
      NumCompAper = vFld(Rs("NCompAper"))
      IdCompAperTrib = vFld(Rs("IdCompAperTrib"))
      NumCompAperTrib = vFld(Rs("NCompAperTrib"))
      If NumCompAper = 0 Then   'parche por si acaso
         NumCompAper = 1
         NumCompAperTrib = 0
      End If
      
      RemIVAUTMAnoAnt = vFld(Rs("RemIVAUTMAnoAnt"))
   End If
   
   Call CloseRs(Rs)

   Lb_Reemp.visible = False
   
   Call SetTxRO(Tx_CompAper, True)

   lIdCompAper = 0

   If IdCompAper = 0 Then
   
      Tx_CompAper = 1    'por ahora no se toma en cuenta el caso de numeración contínua y se parte siempre de 1
              
   Else
      
      lIdCompAper = IdCompAper
      
      If NumCompAper = 0 Then
         'aunque parezca extraño, ocurrió
         Q1 = "SELECT Correlativo "
         Q1 = Q1 & " FROM Comprobante "
         Q1 = Q1 & " WHERE IdComp =" & IdCompAper & " AND Tipo=" & TC_APERTURA & " AND TipoAjuste = " & TAJUSTE_FINANCIERO
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            NumCompAper = vFld(Rs("Correlativo"))
         End If
         
         Call CloseRs(Rs)
         
         Q1 = "UPDATE EmpresasAno SET NCompAper = " & NumCompAper
         Q1 = Q1 & " WHERE idEmpresa=" & lIdEmpresa & " AND Ano=" & lAno
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Tx_CompAper = NumCompAper
      Lb_Reemp.visible = True
      
   End If
   
   lIdCompAperTrib = IdCompAperTrib
   
   lidCuentaResul = 0
   lidCuentaCredIVA = 0
   
   Q1 = "SELECT Tipo, Valor FROM ParamEmpresa WHERE Tipo IN ( 'CTARESULT', 'CTACREDIVA')"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      Select Case vFld(Rs("Tipo"))
      
         Case "CTARESULT"
            lidCuentaResul = vFld(Rs("Valor"))
            lParamCtaResult = True
            
         Case "CTACREDIVA"
            lidCuentaCredIVA = vFld(Rs("Valor"))
            lParamCtaCredIVA = True
            
      End Select
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
         
   If lidCuentaResul = 0 Then
      
      'vemos si está la cuenta de resultado por omisión
      Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas WHERE Codigo='" & COD_CTARESULT & "' AND Descripcion like '*" & DESC_CTARESULT & "*'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then
         lidCuentaResul = vFld(Rs("IdCuenta"))
         Tx_Cuenta = FCase(vFld(Rs("Descripcion"), True))
      End If
   
      Call CloseRs(Rs)
      
   Else
   
      'cargamos la cuenta
      Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas WHERE IdCuenta=" & lidCuentaResul
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then
         Tx_Cuenta = FCase(vFld(Rs("Descripcion"), True))
      End If
   
      Call CloseRs(Rs)
    
   End If
   
   If lidCuentaCredIVA = 0 Then

      'vemos si está la cuenta de credito IVA por omisión
      Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas WHERE Codigo='" & COD_CTACREDIVA & "' AND Descripcion like '*" & DESC_CTACREDIVA & "*'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then
         lidCuentaCredIVA = vFld(Rs("IdCuenta"))
         Tx_CtaCredIVA = FCase(vFld(Rs("Descripcion"), True))
      End If
   
      Call CloseRs(Rs)
      
   Else
   
      'cargamos la cuenta
      Q1 = "SELECT IdCuenta, Descripcion FROM Cuentas WHERE IdCuenta=" & lidCuentaCredIVA
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then
         Tx_CtaCredIVA = FCase(vFld(Rs("Descripcion"), True))
      End If
   
      Call CloseRs(Rs)
      
   End If
   
   
   'Si existe año anterior, bloqueamos ingreso de remanente y
   'Tomamos remanente de IVA que se almacenó al cerrar año anterior, en el registro del año anterior
   If gEmpresa.TieneAnoAnt Then
   
      Q1 = "SELECT RemIVAUTM "
      Q1 = Q1 & " FROM EmpresasAno "
      Q1 = Q1 & " WHERE idEmpresa=" & lIdEmpresa & " AND Ano=" & lAno - 1
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then
         RemIVAUTM = vFld(Rs("RemIVAUTM"))
      End If
   
      Call CloseRs(Rs)

      Tx_RemIVAUTMAnoAnt = Format(RemIVAUTM, DBLFMT2)
   
'      Call SetTxRO(Tx_RemIVAUTMAnoAnt, True)   '7 ago 2020 Victoe Morales indica dejar siempre habilitado
   
   'Si no existe año anterior, tomamos remanente de IVA que se almacenó en apertura anterior de este año, si es que la hubo
   Else
      Tx_RemIVAUTMAnoAnt = Format(RemIVAUTMAnoAnt, DBLFMT2)
   
   End If
   
End Sub
Private Sub Tx_CompAper_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_EMPRESA) Then
      Call EnableForm(Me, False)
   End If
   
End Function

Public Sub SaveCuentas()
   Dim Q1 As String
   Dim Rs As Recordset
   
   If lParamCtaResult Then
      Q1 = "UPDATE ParamEmpresa SET Valor = " & lidCuentaResul & " WHERE Tipo = 'CTARESULT' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Valor, IdEmpresa, Ano) VALUES('CTARESULT', " & lidCuentaResul & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
      
   End If
   
   Call ExecSQL(DbMain, Q1)

   If lParamCtaCredIVA Then
      Q1 = "UPDATE ParamEmpresa SET Valor = " & lidCuentaCredIVA & " WHERE Tipo = 'CTACREDIVA' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Valor, IdEmpresa, Ano) VALUES('CTACREDIVA', " & lidCuentaCredIVA & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
      
   End If
   
   Call ExecSQL(DbMain, Q1)
   gCtasBas.IdCtaCredIVA = lidCuentaCredIVA

   'guardamos remanente de IVA si la empresa no tiene año anterior y, por lo tanto, es de ingreso directo (si no, viene de año anterior)
   Q1 = "UPDATE EmpresasAno SET RemIVAUTMAnoAnt = " & str(vFmt(Tx_RemIVAUTMAnoAnt))
   Q1 = Q1 & " WHERE idEmpresa=" & lIdEmpresa & " AND Ano=" & lAno
   
   Call ExecSQL(DbMain, Q1)
      
End Sub

Private Sub Tx_RemIVAUTMAnoAnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call Tx_RemIVAUTMAnoAnt_LostFocus
      KeyAscii = 0
   Else
      Call KeyDecPos(KeyAscii)
   End If
      
End Sub

Private Sub Tx_RemIVAUTMAnoAnt_LostFocus()
   Tx_RemIVAUTMAnoAnt = Format(vFmt(Tx_RemIVAUTMAnoAnt), DBLFMT2)
End Sub
