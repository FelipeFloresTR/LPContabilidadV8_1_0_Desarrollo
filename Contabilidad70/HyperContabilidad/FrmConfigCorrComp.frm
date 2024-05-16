VERSION 5.00
Begin VB.Form FrmConfigCorrComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Comprobantes"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   Icon            =   "FrmConfigCorrComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Fecha Comprobante Centralización"
      Height          =   1935
      Left            =   6240
      TabIndex        =   35
      Top             =   4020
      Width           =   4935
      Begin VB.OptionButton Op_DtCompCent 
         Caption         =   "Mantener configuración predefinida por el sistema"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   4515
      End
      Begin VB.TextBox Tx_DayDtCompCent 
         Height          =   315
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1380
         Width           =   375
      End
      Begin VB.OptionButton Op_DtCompCent 
         Caption         =   "Asignar día           del mes a la fecha de los comprobantes"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   1380
         Width           =   4515
      End
      Begin VB.OptionButton Op_DtCompCent 
         Caption         =   "Asignar último día del mes a la fecha de los comprobantes"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   4515
      End
      Begin VB.Label Label5 
         Caption         =   " "
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   37
         Top             =   960
         Width           =   2595
      End
      Begin VB.Label Label6 
         Caption         =   "de centralización"
         Height          =   195
         Left            =   420
         TabIndex        =   36
         Top             =   1020
         Width           =   1875
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Impresión de Comprobante"
      Height          =   2055
      Left            =   6240
      TabIndex        =   29
      Top             =   1800
      Width           =   4935
      Begin VB.CheckBox Ch_TituloTipoComp 
         Caption         =   "Incluir Tipo de Comprobante en el Título"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Contenido Columna Detalle Movimiento"
         Height          =   1095
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   4395
         Begin VB.OptionButton Op_PrtMovDet 
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton Op_PrtMovDet 
            Caption         =   "Entidad"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   11
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton Op_PrtMovDet 
            Caption         =   "Centro de Gestión"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1875
         End
         Begin VB.OptionButton Op_PrtMovDet 
            Caption         =   "Área de Negocio"
            Height          =   195
            Index           =   4
            Left            =   2520
            TabIndex        =   13
            Top             =   720
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresión Resumida"
      Height          =   1935
      Left            =   1080
      TabIndex        =   27
      Top             =   4020
      Width           =   4935
      Begin VB.CommandButton Bt_MarcarRes 
         Caption         =   "Aplicar a los comprobantes ya existentes"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1020
         Width           =   3255
      End
      Begin VB.CheckBox Ch_ImpResCent 
         Caption         =   "Todo nuevo comprobante de  centralización queda con la"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Si no marca esta opción, los comprobantes toman el estado Pendiente y deben ser aprobados manualmente."
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Lb_ImpRes 
         Caption         =   "opción 'Imprimir Resumido' automáticamente seleccionada."
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   660
         Width           =   4215
      End
   End
   Begin VB.OptionButton Op_PerCorr 
      Caption         =   "Contínuo (no se reinicia con cada período)"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   26
      Top             =   7620
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Fr_OpComp 
      Caption         =   "Opciones"
      Height          =   1455
      Left            =   6240
      TabIndex        =   24
      Top             =   180
      Width           =   4935
      Begin VB.CheckBox Ch_CompAnuladoLibDiario 
         Caption         =   "Se muestran los comprobantes anulados en el Libro Diario"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Si no marca esta opción, los comprobantes toman el estado Pendiente y deben ser aprobados manualmente."
         Top             =   1020
         Width           =   4515
      End
      Begin VB.CheckBox Ch_AbrirMesesParalelo 
         Caption         =   "Se permite abrir más de un mes en paralelo"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Si no marca esta opción, los comprobantes toman el estado Pendiente y deben ser aprobados manualmente."
         Top             =   720
         Width           =   4515
      End
      Begin VB.CheckBox Ch_CompAprobado 
         Caption         =   "Todo comprobante queda Aprobado al momento de crearlo"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Si no marca esta opción, los comprobantes toman el estado Pendiente y deben ser aprobados manualmente."
         Top             =   420
         Width           =   4515
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   360
      Picture         =   "FrmConfigCorrComp.frx":000C
      ScaleHeight     =   675
      ScaleWidth      =   615
      TabIndex        =   23
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   11400
      TabIndex        =   19
      Top             =   660
      Width           =   1035
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   11400
      TabIndex        =   18
      Top             =   300
      Width           =   1035
   End
   Begin VB.Frame Fr_CorrComp 
      Caption         =   "Correlativo Comprobantes"
      Height          =   3675
      Left            =   1080
      TabIndex        =   20
      Top             =   180
      Width           =   4935
      Begin VB.Frame Fr_TipoComp 
         Caption         =   "Tipo"
         Height          =   1155
         Left            =   180
         TabIndex        =   22
         Top             =   300
         Width           =   4515
         Begin VB.OptionButton Op_TipoCorr 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   720
            Width           =   230
         End
         Begin VB.OptionButton Op_TipoCorr 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Único (independiente del tipo de comprobante)"
            Height          =   195
            Index           =   2
            Left            =   460
            TabIndex        =   34
            Top             =   720
            Width           =   3315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Por tipo de comprobante (Ingreso, Egreso o Traspaso)"
            Height          =   195
            Index           =   1
            Left            =   460
            TabIndex        =   33
            Top             =   360
            Width           =   3825
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periodo"
         Height          =   795
         Left            =   180
         TabIndex        =   21
         Top             =   1920
         Width           =   4515
         Begin VB.OptionButton Op_PerCorr 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   3
            Top             =   360
            Width           =   255
         End
         Begin VB.OptionButton Op_PerCorr 
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   360
            Width           =   230
         End
         Begin VB.Label Label3 
            Caption         =   "Anual"
            Height          =   255
            Index           =   0
            Left            =   3060
            TabIndex        =   32
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Mensual"
            Height          =   255
            Left            =   465
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "La renumeración con ingreso de un correlativo inicial sólo es válida para PERIODO ANUAL"
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   180
         TabIndex        =   25
         Top             =   2940
         Width           =   4335
      End
   End
End
Attribute VB_Name = "FrmConfigCorrComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIndexTipoCorr As Integer
Dim lIndexPerCorr As Integer
Dim lRc As Integer

Public Function FEdit() As Integer
   Me.Show vbModal
   
   FEdit = lRc
   
End Function

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_MarcarRes_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As String
   Dim sFrom As String
   Dim sSet As String
   Dim sWhere As String
   
   Me.MousePointer = vbHourglass
   
   Tbl = "Comprobante"
   sFrom = " Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   sSet = " Comprobante.ImpResumido = 1 "
   sWhere = " WHERE MovComprobante.DeCentraliz <> 0"
   sWhere = sWhere & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
    '3376884
    Call SeguimientoComprobantes(0, gEmpresa.id, gEmpresa.Ano, "FrmConfigCorrComp.Bt_MarcarRes_Click", "", 1, sWhere, gUsuario.IdUsuario, 1, 2)
    'fin 3376884
   
   Me.MousePointer = vbDefault
   
   MsgBox1 "Todos los comprobantes de centralización han quedado con la opción 'Imprimir Resumido' seleccionada.", vbInformation
   
End Sub

Private Sub Form_Load()
   Dim Rs As Recordset
   Dim Q1 As String

   Call EnableForm(Me, gEmpresa.FCierre = 0)
   Op_DtCompCent(DTCOMPCENT_CURRDEF) = 1
   
      
   Call LoadAll
   
   Call SetupPriv
   
   If Not gFunciones.ComprobanteResumido Then
      Ch_ImpResCent.visible = False
      Ch_ImpResCent.Enabled = False
      Lb_ImpRes.visible = False
      Fr_OpComp.Height = Fr_OpComp.Height - 500
      Me.Height = Me.Height - 500
   End If
   
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo <> " & TC_APERTURA               'WHERE Estado <> " & EC_ANULADO & " OR Tipo <> " & TC_APERTURA)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante que no es de apertura
      Fr_CorrComp.Enabled = False
   End If
      
   Call CloseRs(Rs)

      
End Sub

Private Sub Bt_OK_Click()
   
   If Not valida() Then
      Exit Sub
   End If
   
   Call SaveAll
   lRc = vbOK
   
   Unload Me
 
End Sub

Private Sub SaveAll()
   Dim Rc As Long
   Dim Rs As Recordset
   Dim Estado As Integer
   Dim AbrirMesesParalelo As Integer
   Dim CompAnuladoLibDiario As Integer
   Dim ImpRes As Integer, TitTipoComp As Integer
   Dim CompImpEntidad As Integer
   Dim i As Integer
   Dim NewOpt As Integer
   Dim DtCompCent As Integer
   Dim Q1 As String
      
   'Correlativos Comprobantes
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'TCORRCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
         
      'actualizamos
      If gTipoCorrComp <> lIndexTipoCorr Then
         Q1 = "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & lIndexTipoCorr & "' WHERE Tipo = 'TCORRCOMP'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Call ExecSQL(DbMain, Q1)
      End If
      
      If gPerCorrComp <> lIndexPerCorr Then
         Q1 = "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & lIndexPerCorr & "' WHERE Tipo = 'PCORRCOMP'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
      
   Else
   
      'no existe, insertamos
      Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano ) VALUES ('TCORRCOMP', 0, '" & lIndexTipoCorr & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano ) VALUES ('PCORRCOMP', 0, '" & lIndexPerCorr & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      
   End If
   
   Call CloseRs(Rs)
   
   gTipoCorrComp = lIndexTipoCorr
   gPerCorrComp = lIndexPerCorr
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'ESTADOCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Ch_CompAprobado <> Abs(gEstadoNewComp = EC_APROBADO) Then 'cambió
      
      Estado = EC_PENDIENTE
      If Ch_CompAprobado <> 0 Then
         Estado = EC_APROBADO
      End If
      
      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & Estado & "' WHERE Tipo = 'ESTADOCOMP' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano ) VALUES ('ESTADOCOMP', 0, '" & Estado & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If
      
      gEstadoNewComp = Estado
      
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'MESPARALEL'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Ch_AbrirMesesParalelo <> Abs(gAbrirMesesParalelo = True) Then 'cambió

      AbrirMesesParalelo = False
      If Ch_AbrirMesesParalelo <> 0 Then
         AbrirMesesParalelo = True
      End If

      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & AbrirMesesParalelo & "' WHERE Tipo = 'MESPARALEL' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano ) VALUES ('MESPARALEL', 0, '" & AbrirMesesParalelo & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gAbrirMesesParalelo = AbrirMesesParalelo

   End If

   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'VERCOMPANU'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Ch_CompAnuladoLibDiario <> Abs(gCompAnuladoLibDiario = True) Then 'cambió

      CompAnuladoLibDiario = False
      If Ch_CompAnuladoLibDiario <> 0 Then
         CompAnuladoLibDiario = True
      End If

      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & CompAnuladoLibDiario & "' WHERE Tipo = 'VERCOMPANU' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano ) VALUES ('VERCOMPANU', 0, '" & CompAnuladoLibDiario & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gCompAnuladoLibDiario = CompAnuladoLibDiario

   End If

   Call CloseRs(Rs)
   
   
   For i = 1 To MAX_PRTMOV
   
      If Op_PrtMovDet(i) <> 0 Then
         NewOpt = i
         Exit For
      End If
   Next i
   
   If NewOpt <> gPrtMovDetOpt Then 'cambió
   
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PRTMOVDET'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & NewOpt & "' WHERE Tipo = 'PRTMOVDET' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('PRTMOVDET', 0, '" & NewOpt & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gPrtMovDetOpt = NewOpt
      
      Call CloseRs(Rs)

   End If
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'IMPRESCENT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Ch_ImpResCent <> Abs(gImpResCent = True) Then 'cambió

      ImpRes = False
      If Ch_ImpResCent <> 0 Then
         ImpRes = True
      End If

      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & ImpRes & "' WHERE Tipo = 'IMPRESCENT' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('IMPRESCENT', 0, '" & ImpRes & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gImpResCent = ImpRes

   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DTCOMPCENT'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Op_DtCompCent(DTCOMPCENT_CURRDEF) <> 0 Then
      DtCompCent = DTCOMPCENT_CURRDEF
   ElseIf Op_DtCompCent(DTCOMPCENT_LASTDAY) <> 0 Then
      DtCompCent = DTCOMPCENT_LASTDAY
   Else
      DtCompCent = DTCOMPCENT_DEFDAY
   End If
   
   If DtCompCent <> gDtCompCent Or gDayDtCompCent <> vFmt(Tx_DayDtCompCent) Then   'cambió
      
      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & DtCompCent & "' WHERE Tipo = 'DTCOMPCENT' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & vFmt(Tx_DayDtCompCent) & "' WHERE Tipo = 'DYCOMPCENT' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('DTCOMPCENT', 0, '" & DtCompCent & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('DYCOMPCENT', 0, '" & vFmt(Tx_DayDtCompCent) & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gDtCompCent = DtCompCent
      gDayDtCompCent = vFmt(Tx_DayDtCompCent)

   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'TITTIPCOMP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Ch_TituloTipoComp <> Abs(gTituloTipoComp = True) Then 'cambió

      TitTipoComp = False
      If Ch_TituloTipoComp <> 0 Then
         TitTipoComp = True
      End If

      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & TitTipoComp & "' WHERE Tipo = 'TITTIPCOMP' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('TITTIPCOMP', 0, '" & TitTipoComp & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If

      gTituloTipoComp = TitTipoComp

   End If
   
   Call CloseRs(Rs)

   
End Sub

Private Function valida() As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   
   valida = False
   
   'correlativos comprobantes
   If (gTipoCorrComp > 0 Or gPerCorrComp > 0) And (gTipoCorrComp <> lIndexTipoCorr Or gPerCorrComp <> lIndexPerCorr) Then  'ya existe una definición y es distinta a la que había
      'ya existe una definición y es distinta a la que había
      
      Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo <> " & TC_APERTURA
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)       'WHERE Estado <> " & EC_ANULADO & " OR Tipo <> " & TC_APERTURA)
      
      If Rs.EOF = False Then    'hay al menos un comprobante que no es de apertura
      
         MsgBox1 "No es posible cambiar el tipo de correlativo de los comprobantes, hay comprobantes ya ingresados.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If Val(Tx_DayDtCompCent) > 30 Then
      MsgBox1 "Día del mes asignado para fecha de comprobante de centralización inválido. Debe ser menor o igual a 30.", vbExclamation
      Exit Function
   End If
         
   valida = True
End Function

Private Sub LoadAll()
   
   If gTipoCorrComp > 0 Then
      lIndexTipoCorr = gTipoCorrComp
   Else
      lIndexTipoCorr = TCC_UNICO
   End If
   
   Op_TipoCorr(lIndexTipoCorr).Value = True
   
   If gPerCorrComp > 0 Then
      lIndexPerCorr = gPerCorrComp
   Else
      lIndexPerCorr = TCC_ANUAL
   End If
   
   Op_PerCorr(lIndexPerCorr).Value = True
   
   Ch_CompAprobado = Abs(gEstadoNewComp = EC_APROBADO)
   
   Ch_CompAnuladoLibDiario = Abs(gCompAnuladoLibDiario)
   
   If Ch_AbrirMesesParalelo.Enabled = True Then
      Ch_AbrirMesesParalelo = Abs(gAbrirMesesParalelo)
   End If
  
   If gPrtMovDetOpt > 0 And gPrtMovDetOpt <= MAX_PRTMOV Then
      Op_PrtMovDet(gPrtMovDetOpt) = True
   Else
      Op_PrtMovDet(PRTMOV_DESC) = True
   End If
   
   If Ch_ImpResCent.Enabled = True Then
      Ch_ImpResCent = Abs(gImpResCent)
   End If
  
   If gDtCompCent > 0 Then
      Op_DtCompCent(gDtCompCent) = 1
   End If
   If gDayDtCompCent > 0 Then
      Tx_DayDtCompCent = gDayDtCompCent
   End If
  
   Ch_TituloTipoComp = Abs(gTituloTipoComp)
  
End Sub

Private Sub Op_PerCorr_Click(Index As Integer)
   lIndexPerCorr = Index
   
'   If Index = TCC_MENSUAL Then
      Ch_AbrirMesesParalelo.Enabled = True
'   Else
'      Ch_AbrirMesesParalelo = False
'      Ch_AbrirMesesParalelo.Enabled = False
'   End If
   
End Sub

Private Sub Op_TipoCorr_Click(Index As Integer)
   lIndexTipoCorr = Index
End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   
End Function


Private Sub Tx_DayDtCompCent_Change()
   Op_DtCompCent(DTCOMPCENT_DEFDAY) = 1
   
End Sub

Private Sub Tx_DayDtCompCent_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub
