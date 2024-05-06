VERSION 5.00
Begin VB.Form FrmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Empresa"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "FrmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_RecuDocuSql 
      Caption         =   "Recu. Documentos SQL"
      Height          =   855
      Left            =   9120
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recu. Documentos"
      Height          =   855
      Left            =   9000
      TabIndex        =   42
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Bt_CuadraturaDoc 
      Caption         =   "Cuadratura Documentos"
      Height          =   735
      Left            =   9000
      TabIndex        =   41
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Duplicados 
      Caption         =   "Eliminar Documentos Duplicados"
      Height          =   735
      Left            =   9000
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame8 
      Caption         =   "Activo Fijo"
      Height          =   1035
      Left            =   5040
      TabIndex        =   34
      Top             =   6420
      Width           =   3495
      Begin VB.CommandButton Bt_ConfigActFijo 
         Caption         =   "Configurar Activo Fijo"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   420
         Width           =   2715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impuestos"
      Height          =   1395
      Left            =   5040
      TabIndex        =   32
      Top             =   7620
      Width           =   3495
      Begin VB.CommandButton bt_ConfigImp 
         Caption         =   "Configurar Impuestos"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   420
         Width           =   2715
      End
   End
   Begin VB.CommandButton Bt_OK 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9000
      TabIndex        =   20
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame Frame7 
      Caption         =   "Entidades"
      Height          =   1035
      Index           =   1
      Left            =   300
      TabIndex        =   30
      Top             =   6420
      Width           =   4605
      Begin VB.CommandButton Bt_FmtImpEnt 
         Caption         =   "?"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   420
         Width           =   270
      End
      Begin VB.CommandButton Bt_ImportEnt 
         Caption         =   "Importar Entidades"
         Height          =   375
         Left            =   1620
         TabIndex        =   15
         Top             =   420
         Width           =   2475
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   3
         Left            =   300
         Picture         =   "FrmConfig.frx":000C
         ScaleHeight     =   630
         ScaleWidth      =   690
         TabIndex        =   31
         Top             =   300
         Width           =   690
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Comprobantes"
      Height          =   1095
      Index           =   0
      Left            =   300
      TabIndex        =   27
      Top             =   5220
      Width           =   8265
      Begin VB.CommandButton Bt_GenCompAp 
         Caption         =   "Generar Comprobante de Apertura"
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         ToolTipText     =   $"FrmConfig.frx":0687
         Top             =   480
         Width           =   2715
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   1
         Left            =   300
         Picture         =   "FrmConfig.frx":070F
         ScaleHeight     =   675
         ScaleWidth      =   615
         TabIndex        =   28
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton Bt_ConfigCorrComp 
         Caption         =   "Configurar Comprobantes"
         Height          =   375
         Left            =   1620
         TabIndex        =   13
         Top             =   480
         Width           =   2715
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Definición Plan de Cuentas"
      Height          =   4935
      Left            =   300
      TabIndex        =   22
      Top             =   180
      Width           =   8265
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   300
         Picture         =   "FrmConfig.frx":0CDF
         ScaleHeight     =   630
         ScaleWidth      =   600
         TabIndex        =   33
         Top             =   540
         Width           =   600
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Index           =   0
         Left            =   300
         Picture         =   "FrmConfig.frx":1282
         ScaleHeight     =   630
         ScaleWidth      =   690
         TabIndex        =   26
         Top             =   3900
         Width           =   690
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   1380
         TabIndex        =   25
         Top             =   3780
         Width           =   6555
         Begin VB.CommandButton Bt_CtasBasicas 
            Caption         =   "Definir Cuentas Básicas"
            Height          =   375
            Left            =   3600
            TabIndex        =   12
            Top             =   360
            Width           =   2715
         End
         Begin VB.CommandButton Bt_SaldosAp 
            Caption         =   "Ingresar/Listar Saldos de Apertura"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2715
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Planes propios"
         Height          =   3195
         Left            =   4740
         TabIndex        =   24
         Top             =   420
         Width           =   3195
         Begin VB.CommandButton Bt_FmtExpCuentas 
            Caption         =   "?"
            Height          =   375
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2040
            Width           =   270
         End
         Begin VB.CommandButton Bt_ExportPlan 
            Caption         =   "Exportar Plan de Cuentas"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   2040
            Width           =   2475
         End
         Begin VB.CommandButton Bt_FmtImpCuentas 
            Caption         =   "?"
            Height          =   375
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1500
            Width           =   270
         End
         Begin VB.CommandButton Bt_Niveles 
            Caption         =   "Niveles Plan de Cuentas"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   420
            Width           =   2715
         End
         Begin VB.CommandButton Bt_ImportPlan 
            Caption         =   "Importar Plan de Cuentas"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   1500
            Width           =   2475
         End
         Begin VB.CommandButton Bt_CopyPlanEmp 
            Caption         =   "Copiar Plan de Otra Empresa"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   2715
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Planes pre-definidos por el sistema"
         Height          =   3195
         Left            =   1380
         TabIndex        =   23
         Top             =   420
         Width           =   3195
         Begin VB.CommandButton Bt_VerPlan 
            Caption         =   "?"
            Height          =   375
            Index           =   4
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Ver Plan de Cuentas IFRS"
            Top             =   2040
            Width           =   270
         End
         Begin VB.CommandButton Bt_VerPlan 
            Caption         =   "?"
            Height          =   375
            Index           =   3
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ver Plan de Cuentas Avanzado"
            Top             =   1500
            Width           =   270
         End
         Begin VB.CommandButton Bt_VerPlan 
            Caption         =   "?"
            Height          =   375
            Index           =   2
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Ver Plan de Cuentas Intermedio"
            Top             =   960
            Width           =   270
         End
         Begin VB.CommandButton Bt_VerPlan 
            Caption         =   "?"
            Height          =   375
            Index           =   1
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Ver Plan de Cuentas Básico"
            Top             =   420
            Width           =   270
         End
         Begin VB.CommandButton Bt_EditPlan 
            Caption         =   "Ver Plan de Cuentas Actual"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   2580
            Width           =   2715
         End
         Begin VB.CommandButton Bt_PlanPreDef 
            Caption         =   "Utilizar Plan AVANZADO"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   2
            Top             =   1500
            Width           =   2460
         End
         Begin VB.CommandButton Bt_PlanPreDef 
            Caption         =   "Utilizar Plan INTERMEDIO"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   1
            Top             =   960
            Width           =   2460
         End
         Begin VB.CommandButton Bt_PlanPreDef 
            Caption         =   "Utilizar Plan BÁSICO"
            Height          =   375
            Index           =   1
            Left            =   225
            TabIndex        =   0
            Top             =   420
            Width           =   2460
         End
         Begin VB.CommandButton Bt_PlanPreDef 
            Caption         =   "Utilizar Plan IFRS"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   3
            Top             =   2040
            Width           =   2460
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informes (opcional)"
      Height          =   1395
      Left            =   300
      TabIndex        =   21
      Top             =   7620
      Width           =   4605
      Begin VB.CommandButton Bt_Firmas 
         Caption         =   "Configurar Firmas"
         Height          =   375
         Left            =   1620
         TabIndex        =   39
         Top             =   960
         Width           =   2715
      End
      Begin VB.CommandButton bt_Opciones 
         Caption         =   "Configurar Informes"
         Height          =   375
         Left            =   1620
         TabIndex        =   18
         Top             =   420
         Width           =   2715
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Index           =   2
         Left            =   300
         Picture         =   "FrmConfig.frx":190D
         ScaleHeight     =   705
         ScaleWidth      =   675
         TabIndex        =   29
         Top             =   300
         Width           =   675
      End
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PLAN_BASICO = 1
Const PLAN_INTERMEDIO = 2
Const PLAN_AVANZADO = 3
Const PLAN_IFRS = 4



Private Sub Bt_ConfigActFijo_Click()
   Dim Frm As FrmConfigActFijo
   
   Set Frm = New FrmConfigActFijo
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_ConfigCorrComp_Click()
   Dim Frm As FrmConfigCorrComp
   
   Set Frm = New FrmConfigCorrComp
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub bt_ConfigImp_Click()
   Dim Frm As FrmIVA
   
   Set Frm = New FrmIVA
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_CopyPlanEmp_Click()
   Dim Frm As FrmCopyPlan
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Rc As Integer
   
   'verificamos que no hayan comprobantes ingresados, que tengan movimientos
   'si hay comprobante de apertura vacío, que se genera automáticamente para guardar número, no damos mensaje
   Q1 = "SELECT Comprobante.IdComp, Tipo FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay comprobantes ya ingresados.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)
   
   'verificamos que no hayan documentos ingresados, que tengan movimientos con cuentas asociadas
   Q1 = "SELECT IdMovDoc FROM MovDocumento WHERE IdCuenta <> 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay documentos ya ingresados que hacen referencia a cuentas de este plan.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)

   
   Set Frm = New FrmCopyPlan
   Rc = Frm.FCopy()
   Set Frm = Nothing
   
End Sub

Private Sub Bt_CtasBasicas_Click()
   Dim Frm As FrmConfigCtasDef
   
   Set Frm = New FrmConfigCtasDef
   Frm.Show vbModal
   Set Frm = Nothing

End Sub

Private Sub Bt_CuadraturaDoc_Click()
Me.MousePointer = vbHourglass

        Dim Q1 As String

        Dim PathDbAnoAnt As String
        Dim ConnStr As String

        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If

   If gDbType = SQL_ACCESS Then
        PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")

        If ExistFile(PathDbAnoAnt) Then
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)

        Else
         MsgBox1 "No se encontró la base de datos del año anterior. No es posible generar Cuadratura de documentos.", vbExclamation + vbOKOnly
         Me.MousePointer = vbDefault
          Exit Sub
        End If
    End If

             Q1 = ""
             Q1 = "Update Documento Set FExported = 0 "
             Q1 = Q1 & " Where IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
             Q1 = Q1 & " And (FExported <> 0 or FExported is not null) "

             Call ExecSQL(DbAnoAnt, Q1)
              If gDbType = SQL_ACCESS Then
             Call CloseDb(DbAnoAnt)
             End If
                     
     Call GenDocsPendientes(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, False, True)
    'Call ComprobanteApeturaFexported
     If gDbType = SQL_ACCESS Then
            'Call CorrigeDuplicados(False)
           ' Call CorrigePagadosAñoAnteriores(False)
     End If
Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
MsgBox1 "Cuadratura de Documentos Terminado.", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
End Sub

'14202137
Private Sub Bt_Duplicados_Click()
Me.MousePointer = vbHourglass

If gDbType = SQL_ACCESS Then
 'Call CorrigeDuplicados(True)
'Call CorrigePagadosAñoAnteriores(True)
 Call CorrigePagadosAñoAnterioresPrueba(False)
 Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
 End If
Me.MousePointer = vbDefault
End Sub
'14202137

Private Sub Bt_EditPlan_Click()
   Dim Frm As FrmPlanCuentas
   
   Set Frm = New FrmPlanCuentas
   Call Frm.FEdit
   Set Frm = Nothing

End Sub
Private Sub Bt_ExportPlan_Click()
   Dim Rc As Integer
   Dim ExistePlan As Boolean
   Dim Dig(MAX_NIVELES) As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   If MsgBox1("Para realizar esta operación, nadie debe estar modificando el plan de cuentas de esta empresa, para este año." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   'Chequeamos que haya plan de cuentas definido, si no es así, se da un warning porque no se exportará
   Q1 = "SELECT Count(*) as Cant FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExistePlan = True
      End If
   End If
   Call CloseRs(Rs)
   
   If Not ExistePlan Then
      MsgBox1 "El plan de cuentas está vacío.", vbInformation + vbOKOnly
      Exit Sub
   End If
   
   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = "PlanCtas-" & gEmpresa.Rut & ".txt"
   FrmMain.Cm_ComDlg.InitDir = gExportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Exportación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowSave
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   MousePointer = vbHourglass
   DoEvents
      
   Rc = ExportarCuentas(FrmMain.Cm_ComDlg.Filename, gNiveles.nNiveles, gNiveles.Largo()) > 0
   
   MousePointer = vbDefault

End Sub

Private Sub Bt_Firmas_Click()
Dim Frm As FrmConfigInformeFirma
   
   Set Frm = New FrmConfigInformeFirma
   Frm.Show vbModal
   Set Frm = Nothing
End Sub

Private Sub Bt_FmtExpCuentas_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewCuentas
   Set Frm = Nothing

End Sub

Private Sub Bt_FmtImpCuentas_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewCuentas
   Set Frm = Nothing

End Sub

Private Sub Bt_FmtImpEnt_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewEntidad
   Set Frm = Nothing
   
End Sub

Private Sub Bt_GenCompAp_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim InitAno As String
   Dim IdCompAper As Long
   Dim Rc As Integer
   Dim HayComp As Boolean
   Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim ConnStr As String
    
   If Not ValidaIngresoComp(True) Then
      Exit Sub
   End If
      
   Me.MousePointer = vbHourglass
      
   'veamos si la empresa tiene historia (año anterior a partir del cual se generó este año)
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo='INITAÑO' "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then 'está
      InitAno = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
         
   'veamos si ya se generó comprobante de apertura
   Q1 = "SELECT IdCompAper"
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      IdCompAper = vFld(Rs("IdCompAper"))
   End If
   
   Call CloseRs(Rs)
   
   If IdCompAper = 0 Then
   
      ' Veamos si exsite un comprobante de aprtura generado manualmente, dado que no está registrado en la tabla EmpresasAño
      Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdCompAper = vFld(Rs("IdComp"))
      End If
      
      Call CloseRs(Rs)
   
   End If
   
  
   'si ya existe un comp de apertura, lo regeneramos
   If IdCompAper <> 0 Then

      'veamos si el ID del comprobante deapertura corresponde al almacenado en la tabla EmpresasAño de la LPContab
      Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         
         If IdCompAper <> vFld(Rs("IdComp")) Then   'no calza el IdCompAper de EmpresasAño con el comprobante en la base del año
            
            IdCompAper = 0
            
            Q1 = "UPDATE EmpresasAno SET IdCompAper = 0 "
            Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      Else
         IdCompAper = 0
                  
      End If
      
      Call CloseRs(Rs)
      
      Rc = GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
   
   Else        'IdCompAper = 0
           
      Q1 = "SELECT IdComp, Fecha FROM Comprobante "
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY Fecha DESC, IdComp DESC"
      Set Rs = OpenRs(DbMain, Q1)
      
      HayComp = (Not Rs.EOF)
      
      Call CloseRs(Rs)
   
      If Not HayComp Then 'no hay comprobantes aún
         Rc = GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
      
      ElseIf gTipoCorrComp = TCC_TIPOCOMP Then 'hay comprobantes pero el correlativo es por tipo => podemos generar un comp de apertura con N° 1
         Rc = GenCompApertura(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, InitAno = "EMPHISTORIA")
      
      Else
         MsgBox1 "No es posible generar automáticamente el comprobante de apertura, dado que ya hay comprobantes ingresados y éste debe ser el primero.", vbExclamation
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      
   End If
   
   Me.MousePointer = vbDefault
      
   If Rc Then
      MsgBox1 "El Comprobante de apertura ha sido generado.", vbInformation + vbOKOnly
      
      Me.MousePointer = vbHourglass
      
      If InitAno = "EMPHISTORIA" Then
         
         '14520904
'         Call CorregirSaldosAnoAnterior
         '14520904
         
         Call GenDocsPendientes(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True)
         'Call GenDocsFullPendientes(gEmpresa.Id, gEmpresa.Rut, gEmpresa.Ano, True)
         
         'inicio descomentar en caso de no pasar documentos pensientes en sql server
          '3026009
          'Call GenDocsPendientesEmpJuntas(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True)
          ' SF 13828558 se agrega el ultimo true, para obtener documentos ODF de año anterior
       ' Call GenDocsPendientesEmpJuntas(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True, True)
         'end
         '3026009
         
         Call GenActFijoResidual(gEmpresa.id, gEmpresa.Rut, gEmpresa.Ano, True)
      ElseIf InitAno = "EMPHISTACC" Then
        
        PathDbAnoAnt = Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab")
        
        If ExistFile(PathDbAnoAnt) Then
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)
         
          'ahora obtenemos los documentos centralizados y pagados con saldo pendiente desde el año anterior
          
          Call CopyDocsFromAccessToSQLServerNew(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
          
          'Luego los activos fijos con valor libro mayor que cero o no depreciables del año anteriro
          
          'Call CopyActFijoFromAccessToSQLServer(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
          
          'finalmente generamos los saldos de apertura en el plan de cuentas
          
          Call GenSaldosAperturaAccessFromSQLServer(DbAnoAnt, gEmpresa.id, gEmpresa.Ano)
        End If
      End If
      
'    If gDbType = SQL_ACCESS Then
'        Call CorrigeDuplicados(False)
'        'Call CorrigePagadosAñoAnteriores(False)
'        Call CorrigePagadosAñoAnteriores2(False)
'        Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
'    End If
      
      Me.MousePointer = vbDefault
   End If
   
End Sub

Private Sub Bt_ImportEnt_Click()
   Dim Rc As Integer
   Dim i As Integer
      
   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   MousePointer = vbHourglass
   DoEvents
      
   Rc = ImportarEntidades(FrmMain.Cm_ComDlg.Filename)
   
'   MsgBox1 "La importación de entidades ha finalizado.", vbInformation + vbOKOnly
           
   MousePointer = vbDefault
   

End Sub

Private Sub Bt_Niveles_Click()
   Dim Frm As FrmNiveles
   
   Set Frm = New FrmNiveles
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()
      
   Unload Me
 
End Sub

Private Sub Bt_ImportPlan_Click()
   Dim Rc As Integer
   Dim ExistePlan As Boolean
   Dim Dig(MAX_NIVELES) As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   If MsgBox1("Para realizar esta operación, nadie debe estar modificando el plan de cuentas de esta empresa, para este año." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   'verificamos que no hayan comprobantes ingresados, que tengan movimientos
   'si hay comprobante de apertura vacío, que se genera automáticamente para guardar número, no damos mensaje
   Q1 = "SELECT Comprobante.IdComp, Tipo FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay comprobantes ya ingresados.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)
   
   'verificamos que no hayan documentos ingresados, que tengan movimientos con cuentas asociadas
   Q1 = "SELECT IdMovDoc FROM MovDocumento WHERE IdCuenta <> 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay documentos ya ingresados que hacen referencia a cuentas de este plan.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)

   'Chequeamos que no haya plan de cuentas definido, si es así, se da un warning porque lo perderá
   Q1 = "SELECT Count(*) as Cant FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExistePlan = True
      End If
   End If
   Call CloseRs(Rs)
   
   If ExistePlan Then
      Rc = MsgBox1("Al importar un Plan de Cuentas perderá el Plan de Cuentas actual." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion)
   
      If Rc = vbNo Then
         Exit Sub
      End If
   End If
   
   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   MousePointer = vbHourglass
   DoEvents
   
   'eliminamos las cuentas básicas de ParamEmpresa porque no van a servir para el nuevo plan porque los IDs son distintos
'   Q1 = "DELETE * FROM ParamEmpresa "
'   Q1 = Q1 & " WHERE Left(Tipo,3)='CTA'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Rc = ExecSQL(DbMain, Q1)

   Q1 = " WHERE Left(Tipo,3)='CTA'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "ParamEmpresa", Q1)
   
   'eliminamos las cuentas básicas de la tabla CuentasBasicas porque no van a servir para el nuevo plan porque los IDs son distintos
'   Q1 = "DELETE * FROM CuentasBasicas "
'   Rc = ExecSQL(DbMain, Q1)

   Q1 = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "CuentasBasicas", Q1)
      
   Rc = ImportarCuentas(FrmMain.Cm_ComDlg.Filename, gNiveles.nNiveles, gNiveles.Largo()) > 0
   
   If Rc >= 0 Then
      Call ReadEmpresa
   
      Call Bt_CtasBasicas_Click
   End If
   
   MousePointer = vbDefault
         
End Sub


Private Sub Bt_Opciones_Click()
   Dim Frm As FrmConfigInformes
   
   Set Frm = New FrmConfigInformes
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_PlanPreDef_Click(Index As Integer)
   
   Call GetPlanPreDef(Index)
   
End Sub

Private Sub Bt_RecuDocuSql_Click()
#If DATACON <> 1 Then
    Call GetDocumentosEliminados
#End If
End Sub

Private Sub Bt_SaldosAp_Click()
   Dim Frm As FrmSaldoApertura
   
   Set Frm = New FrmSaldoApertura
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub



Private Sub Bt_VerPlan_Click(Index As Integer)

   Call VerPlanCuentas(Index)
   
End Sub

#If DATACON = 1 Then
Private Sub Command1_Click()
Call CorrigeDocEliminados
End Sub
#End If


Private Sub Form_Load()
      
   'Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   'Actualizamos la configuración por si alguien la cambió
   
   Me.MousePointer = vbHourglass
   Call ReadEmpresa
   Me.MousePointer = vbDefault
   
   Bt_ExportPlan.visible = gFunciones.ExpPlanCuentas
   Bt_FmtExpCuentas.visible = gFunciones.ExpPlanCuentas
   
   Call SetupPriv
   
    If gDbType = SQL_ACCESS Then
      'Bt_Duplicados.visible = True
      'Bt_CuadraturaDoc.visible = True
      'Bt_RecuDocuSql.visible = False
    Else
       'Bt_Duplicados.visible = False
       'Bt_CuadraturaDoc.visible = True
       'Bt_RecuDocuSql.visible = True
       Call GetBotonRecuSQL
       
    End If
            
    
            
End Sub

Private Sub GetPlanPreDef(ByVal Tipo As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim ExistePlan As Boolean
   Dim Rc As Integer
   Dim Plan As String
   Dim Frm As Form
   Dim FrmIFRS As FrmConfigCodIFRS
   Dim AtribLst As String, FldLst As String
   Dim i As Integer
   
   'verificamos que no hayan comprobantes ingresados, que tengan movimientos
   'si hay comprobante de apertura vacío, que se genera automáticamente para guardar número, no damos mensaje
   Q1 = "SELECT Comprobante.IdComp, Tipo FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay comprobantes ya ingresados.", vbExclamation + vbOKOnly
      
      If Tipo = PLAN_IFRS Then
         MsgBox1 "El Plan de Cuenta IFRS es solo para empresas que recién se crean y funcionan en el sistema.", vbExclamation + vbOKOnly
      End If
      
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)

   'verificamos que no hayan documentos ingresados, que tengan movimientos con cuentas asociadas
   Q1 = "SELECT IdMovDoc FROM MovDocumento WHERE IdCuenta <> 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then    'hay al menos un comprobante
      
      MsgBox1 "No es posible cambiar el plan de la empresa, hay documentos ya ingresados que hacen referencia a cuentas de este plan.", vbExclamation + vbOKOnly
      
      If Tipo = PLAN_IFRS Then
         MsgBox1 "El Plan de Cuenta IFRS es solo para empresas que recién se crean y funcionan en el sistema.", vbExclamation + vbOKOnly
      End If
      
      Call CloseRs(Rs)
      Exit Sub
            
   End If
   
   Call CloseRs(Rs)


   'Chequeamos que no haya plan de cuentas definido, si es así, se da un warning porque lo perderá
   Q1 = "SELECT Count(*) as Cant FROM Cuentas"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExistePlan = True
      End If
   End If
   Call CloseRs(Rs)
   
   If ExistePlan Then
      Rc = MsgBox1("Al importar un Plan de Cuentas perderá el Plan de Cuentas actual." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion)
   
      If Rc = vbNo Then
         Exit Sub
      End If
   End If

   Me.MousePointer = vbHourglass

'   Q1 = "DELETE * FROM Cuentas"
'   Call ExecSQL(DbMain, Q1)

   Q1 = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "Cuentas", Q1)

   For i = 1 To MAX_ATRIB
      AtribLst = AtribLst & ", Atrib" & i
   Next i
     
   FldLst = ""
   If Not DB_MSSQL Then
      FldLst = "IdCuenta, "
   End If
   
   FldLst = FldLst & "idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22 " & AtribLst & ", CodIFRS_EstRes, CodIFRS_EstFin, CodIFRS, TipoPartida, CodCtaPlanSII "
   Q1 = "INSERT INTO Cuentas (" & FldLst & ", IdEmpresa, Ano ) SELECT " & FldLst & ", " & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano FROM "

   If Tipo = PLAN_BASICO Then
'      Q1 = "INSERT INTO Cuentas SELECT * FROM PlanBasico"
      Q1 = Q1 & " PlanBasico"
      Plan = "BÁSICO"
      
   ElseIf Tipo = PLAN_INTERMEDIO Then
'      Q1 = "INSERT INTO Cuentas SELECT * FROM PlanIntermedio"
      Q1 = Q1 & " PlanIntermedio"
      Plan = "INTERMEDIO"
      
   ElseIf Tipo = PLAN_AVANZADO Then
'      Q1 = "INSERT INTO Cuentas SELECT * FROM PlanAvanzado"
      Q1 = Q1 & " PlanAvanzado"
      Plan = "AVANZADO"
      
   ElseIf Tipo = PLAN_IFRS Then
   
      FldLst = ""
      If Not DB_MSSQL Then
         FldLst = "IdCuenta, "
      End If
      
'      FldLst = "IdCuenta, IdPadre, Codigo, Nombre, Descripcion, Nivel, Estado, Clasificacion, Debe, Haber, TipoCapPropio, CodF22 " & AtribLst & ", Codigo as CodIFRS"
      FldLst = FldLst & "IdPadre, Codigo, Nombre, Descripcion, Nivel, Estado, Clasificacion, Debe, Haber, TipoCapPropio, CodF22 " & AtribLst
   
'      Q1 = "INSERT INTO Cuentas SELECT " & FldLst
'      Q1 = Q1 & " FROM IFRS_PlanIFRS"
      
      Q1 = "INSERT INTO Cuentas (" & FldLst & ", CodIFRS, IdEmpresa, Ano ) SELECT " & FldLst & ", Codigo as CodIFRS, " & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano FROM IFRS_PlanIFRS "
      
      Plan = "IFRS"
   End If
   
   Call ExecSQL(DbMain, Q1)
   
   If DB_MSSQL Then
      Call HilarPadresPlanCuentasPreDef   'ésto se requiere para el caso de empresas juntas en que el IdCuenta es Identity no parte de 1, por lo que el IdPadre se desordena
   End If
      
   'guardamos el tipo de plan seleccionado
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then    'ya existe, lo actualizamos
      Q1 = "UPDATE ParamEmpresa SET Valor = '" & Plan & "' WHERE Tipo = 'PLANCTAS'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES( 'PLANCTAS', 0, '" & Plan & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
   End If
   Call ExecSQL(DbMain, Q1)
   Call CloseRs(Rs)
   
   If (Plan = "BÁSICO" Or Plan = "INTERMEDIO" Or Plan = "AVANZADO") And gEmpresa.Ano >= 2017 Then    'se eliminan códigos F22 ya no válidos desde 2017
      Q1 = "UPDATE Cuentas SET CodF22 = '' WHERE CodF22 NOT IN (" & LSTCODF22_2017 & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
   End If
   
   gPlanCuentas = Plan
   
   'ajustamos los niveles del plan por si el usuario los cambió
   Q1 = "UPDATE ParamEmpresa SET Valor=1 WHERE Tipo='DIGNIV1'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ParamEmpresa SET Valor=2 WHERE Tipo='DIGNIV2'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ParamEmpresa SET Valor=2 WHERE Tipo='DIGNIV3'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ParamEmpresa SET Valor=2 WHERE Tipo='DIGNIV4'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ParamEmpresa SET Valor=0 WHERE Tipo='DIGNIV5'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE ParamEmpresa SET Valor=4 WHERE Tipo='NIVELES'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Call ConfigNiveles

   If Tipo <> PLAN_IFRS Then

      'corregimos Comprobantes Tipo para planes Intermedio y Basico
      Call UpdateComprobantesTipo
   End If
   
   
   'seteamos las cuentas básicas
   Call SetCtasBasDef
   
   Me.MousePointer = vbDefault

   MsgBox "El Plan " & Plan & " ha sido cargado con éxito.", vbInformation + vbOKOnly
   MsgBox "Verifique las cuentas básicas definidas para este plan, utilizando el botón ""Definición Cuentas Básicas""", vbInformation + vbOKOnly

   Set Frm = New FrmConfigCtasDef
   Frm.Show vbModal
   Set Frm = Nothing
   
   MsgBox "Verifique la configuración de cuentas para los informes IFRS utilizando la opción" & vbCrLf & vbCrLf & """Definiciones >> Plan de Cuentas >> Configuración Códigos IFRS""", vbInformation + vbOKOnly
   
   Me.MousePointer = vbHourglass
   
   Set FrmIFRS = New FrmConfigCodIFRS
   FrmIFRS.Show vbModal
   Set FrmIFRS = Nothing

   Me.MousePointer = vbDefault
   
End Sub
Private Sub SetCtasBasDef()
   Dim Rs As Recordset
   Dim Q1 As String

   'limpiamos las cuentas
   Call CleanCtasBas(gCtasBas)
   
   'obtenemos los Ids de las cuentas de acuerdo al nombre corto
   
   'IVA Crédito
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='IVACRE' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaIVACred = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaIVACred <> 0 Then
         Call UpdParamEmpresa("CTAIVACRED", 0, gCtasBas.IdCtaIVACred)
      End If
   End If
   Call CloseRs(Rs)
   
   'IVA Débito
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='IVADEB' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaIVADeb = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaIVADeb <> 0 Then
         Call UpdParamEmpresa("CTAIVADEB", 0, gCtasBas.IdCtaIVADeb)
      End If
   End If
   Call CloseRs(Rs)
   
   'Otros Impuestos Crédito
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='OIMPRE' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaOtrosImpCred = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaOtrosImpCred <> 0 Then
         Call UpdParamEmpresa("CTAOIMPCRE", 0, gCtasBas.IdCtaOtrosImpCred)
      End If
   End If
   Call CloseRs(Rs)
   
   'Otros Impuestos Débito
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='OIMPPA' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaOtrosImpDeb = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaOtrosImpDeb <> 0 Then
         Call UpdParamEmpresa("CTAOIMPDEB", 0, gCtasBas.IdCtaOtrosImpDeb)
      End If
   End If
   Call CloseRs(Rs)
   
   'Impuesto único a los trabajadores
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='IMPUNICO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaImpUnico = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaImpUnico <> 0 Then
         Call UpdParamEmpresa("CTAIMPUNIC", 0, gCtasBas.IdCtaImpUnico)
      End If
   End If
   Call CloseRs(Rs)
   
   'Pago facturas
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='BANCO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaPagoFacturas = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaPagoFacturas <> 0 Then
         Call UpdParamEmpresa("CTAPAGOFAC", 0, gCtasBas.IdCtaPagoFacturas)
      End If
   End If
   Call CloseRs(Rs)
   
   'Cobro facturas
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='BANCO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaCobFacturas = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaCobFacturas <> 0 Then
         Call UpdParamEmpresa("CTACOBFAC", 0, gCtasBas.IdCtaCobFacturas)
      End If
   End If
   Call CloseRs(Rs)
   
   'Impuesto Retenido
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='IMP2DA' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaImpRet = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaImpRet <> 0 Then
         Call UpdParamEmpresa("CTAIMPRET", 0, gCtasBas.IdCtaImpRet)
      End If
   End If
   Call CloseRs(Rs)
   
   'Honorarios por pagar
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='HONPAG' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaNetoHon = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaNetoHon <> 0 Then
         Call UpdParamEmpresa("CTANETORET", 0, gCtasBas.IdCtaNetoHon)
      End If
   End If
   Call CloseRs(Rs)
   
   'Dieta directorio
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='DIETA' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaNetoDieta = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaNetoDieta <> 0 Then
         Call UpdParamEmpresa("CTANETODIE", 0, gCtasBas.IdCtaNetoDieta)
      End If
   End If
   Call CloseRs(Rs)
   
   'Resultado del ejercicio: Patrimonio
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='PATRIM' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaPatrimonio = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaPatrimonio <> 0 Then
         Call UpdParamEmpresa("CTAPATRIM", 0, gCtasBas.IdCtaPatrimonio)
      End If
   End If
   Call CloseRs(Rs)
   
   'Cuenta de resultado ejercicio
   Set Rs = OpenRs(DbMain, "SELECT IdCuenta FROM Cuentas WHERE Nombre='RESEJE' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Rs.EOF = False Then
      gCtasBas.IdCtaResEje = vFld(Rs("IdCuenta"))
      If gCtasBas.IdCtaResEje <> 0 Then
         Call UpdParamEmpresa("CTARESEJE", 0, gCtasBas.IdCtaResEje)
      End If
   End If
   Call CloseRs(Rs)
   
   'actualizamos plan de cuentas con CodF29 de impuesto único
   Q1 = "UPDATE Cuentas SET CodF29=" & CODF29_IMPUNICO & " WHERE IdCuenta = " & gCtasBas.IdCtaImpUnico
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Call CloseRs(Rs)
      
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Bt_PlanPreDef(1).Enabled = False
      Bt_PlanPreDef(2).Enabled = False
      Bt_PlanPreDef(3).Enabled = False
      Bt_CopyPlanEmp.Enabled = False
      Bt_ImportPlan.Enabled = False
      Bt_ExportPlan.Enabled = False
      Bt_ImportEnt.Enabled = False
      Bt_GenCompAp.Enabled = False
   End If
   
End Function
Private Sub VerPlanCuentas(ByVal IdxPlanCuentas)
   Dim Frm As FrmPlanCuentas
   Dim PlanCuentas As String, NombrePlan As String
   
   Set Frm = New FrmPlanCuentas
   Select Case IdxPlanCuentas
      Case PLAN_BASICO
         PlanCuentas = "PlanBasico"
         NombrePlan = "Básico"
      Case PLAN_INTERMEDIO
         PlanCuentas = "PlanInterMedio"
         NombrePlan = "Intermedio"
      Case PLAN_AVANZADO
         PlanCuentas = "PlanAvanzado"
         NombrePlan = "Avanzado"
      Case PLAN_IFRS
         PlanCuentas = "IFRS_PlanIFRS"
         NombrePlan = "IFRS"
   End Select
         
   Call Frm.FViewPlan(PlanCuentas, NombrePlan)
   Set Frm = Nothing

End Sub

#If DATACON <> 1 Then
'3405779
'Private Sub GetDocumentosEliminados() ' esta funcion recupera los documentos eliminados pagados en año anterior
'   Dim Rs As Recordset
'   Dim Q1 As String
'   Dim Rc As Long
'
'   Me.MousePointer = vbHourglass
'
'    Q1 = ""
'    Q1 = Q1 & " INSERT INTO dbo.MovDocumento"
'    Q1 = Q1 & " (IdEmpresa,Ano,IdDoc,IdCompCent,IdCompPago,Orden,IdCuenta,Debe,Haber,Glosa,"
'    Q1 = Q1 & " IdTipoValLib,EsTotalDoc,IdCCosto,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld)"
'    Q1 = Q1 & " SELECT MV.IdEmpresa,YEAR(DOC.FEmisionOri -2) Ano,DOC.OldIdDoc"
'    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV2 WHERE MV2.IdDoc = DOC.OldIdDoc AND MV2.DeCentraliz = 1 AND MV2.Ano = YEAR(DOC.FEmisionOri - 2) AND MV2.IdEmpresa = DOC.IdEmpresa),0) IdCompCent"
'    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV1 WHERE MV1.IdDoc = DOC.OldIdDoc AND MV1.DePago = 1 AND MV1.Ano = YEAR(DOC.FEmisionOri - 2) AND MV1.IdEmpresa = DOC.IdEmpresa Order by MV1.IdMov desc),0) IdCompPago"
'    Q1 = Q1 & " ,Orden"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = MV.IdCuenta AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(DOC.FEmisionOri - 2) AND CS.IdEmpresa = DOC.IdEmpresa),0) IdCuenta"
'    Q1 = Q1 & " ,Debe,Haber,Glosa,IdTipoValLib,EsTotalDoc,IdCCosto"
'    Q1 = Q1 & " ,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld"
'    Q1 = Q1 & " FROM MovDocumento MV INNER JOIN"
'    Q1 = Q1 & " (SELECT D.IdDoc, D.FEmisionOri,D.OldIdDoc, D.IdEmpresa"
'    Q1 = Q1 & " FROM Documento D"
'    Q1 = Q1 & " LEFT JOIN Documento DO ON DO.IdEmpresa = D.IdEmpresa AND DO.Ano = " & gEmpresa.Ano
'    Q1 = Q1 & " AND DO.NumDoc = D.NumDoc AND DO.TipoLib = D.TipoLib AND DO.TipoDoc = D.TipoDoc"
'    Q1 = Q1 & " WHERE D.Ano = " & gEmpresa.Ano + 1
'    Q1 = Q1 & " AND D.IdEmpresa = " & gEmpresa.id
'    Q1 = Q1 & " AND D.FEmisionOri IS NOT NULL AND D.OldIdDoc IS not NULL AND YEAR(D.FEmisionOri -2) = " & gEmpresa.Ano
'    Q1 = Q1 & " AND DO.IdDoc IS NULL) DOC ON DOC.IdDoc = MV.IdDoc"
'
'    Rc = ExecSQL(DbMain, Q1)
'    'Fin insert mov documento
'
'    If Rc > 0 Then
'
'     Rc = 0
'    '........................
'    Q1 = ""
'    Q1 = " SET IDENTITY_INSERT Documento ON"
'    Q1 = Q1 & " INSERT INTO Documento"
'    Q1 = Q1 & " (IdDoc,IdEmpresa,Ano,IdCompCent,IdCompPago,TipoLib,TipoDoc,NumDoc,NumDocHasta"
'    Q1 = Q1 & " ,IdEntidad,TipoEntidad,RutEntidad,NombreEntidad,FEmision,FVenc,Descrip,Estado,Exento"
'    Q1 = Q1 & " ,IdCuentaExento,Afecto,IdCuentaAfecto,IVA,IdCuentaIVA,OtroImp,IdCuentaOtroImp,Total"
'    Q1 = Q1 & " ,IdCuentaTotal,IdUsuario,FechaCreacion,FEmisionOri,CorrInterno,SaldoDoc,FExported"
'    Q1 = Q1 & " ,OldIdDoc,DTE,PorcentRetencion,TipoRetencion,MovEdited,OtrosVal,FImporF29,NumDocRef"
'    Q1 = Q1 & " ,IdCtaBanco,TipoRelEnt,IdSucursal,TotPagadoAnoAnt,FImportSuc,Giro,FacCompraRetParcial"
'    Q1 = Q1 & " ,IVAIrrecuperable,DocOtrosEnAnalitico,OldIdDocTmp,NumFiscImpr,NumInformeZ,CantBoletas"
'    Q1 = Q1 & " ,VentasAcumInfZ,IdDocAsoc,PropIVA,ValIVAIrrec,IVAInmueble,FImpFacturacion,CodSIIDTEIVAIrrec"
'    Q1 = Q1 & " ,TipoDocAsoc,IVAActFijo,EntRelacionada,NumCuotas,CompraBienRaiz,NumDocAsoc,DTEDocAsoc"
'    Q1 = Q1 & " ,IdANegCCosto,UrlDTE,CodCtaAfectoOld,CodCtaExentoOld,CodCtaTotalOld,DocOtroEsCargo"
'    Q1 = Q1 & " ,ValRet3Porc,IdCuentaRet3Porc,Tratamiento)"
'    Q1 = Q1 & " SELECT D.OldIdDoc IdDoc,D.IdEmpresa,YEAR(D.FEmisionOri - 2) AS Ano"
'    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV WHERE MV.IdDoc = D.OldIdDoc AND MV.DeCentraliz = 1 AND MV.Ano = YEAR(D.FEmisionOri -2) AND MV.IdEmpresa = D.IdEmpresa),0) IdCompCent"
'    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV1 WHERE MV1.IdDoc = D.OldIdDoc AND MV1.DePago = 1 AND MV1.Ano = YEAR(D.FEmisionOri -2) AND MV1.IdEmpresa = D.IdEmpresa order by MV1.IdMov desc),0) IdCompPago"
'    Q1 = Q1 & " ,D.TipoLib,D.TipoDoc,D.NumDoc,D.NumDocHasta,D.IdEntidad,D.TipoEntidad"
'    Q1 = Q1 & " ,D.RutEntidad,D.NombreEntidad,D.FEmision,D.FVenc,D.Descrip,3 Estado,D.Exento"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaExento AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaExento"
'    Q1 = Q1 & " ,D.Afecto"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaAfecto AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaAfecto"
'    Q1 = Q1 & " ,D.IVA"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaIVA AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaIVA"
'    Q1 = Q1 & " ,D.OtroImp"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaOtroImp AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaOtroImp"
'    Q1 = Q1 & " ,D.Total"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaTotal AND CUE.Codigo = CS.Codigo"
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaTotal"
'    Q1 = Q1 & " ,D.IdUsuario,D.FechaCreacion,D.FEmisionOri,D.CorrInterno,D.SaldoDoc,D.FExported,0 as OldIdDoc"
'    Q1 = Q1 & " ,D.DTE,D.PorcentRetencion,D.TipoRetencion,D.MovEdited,D.OtrosVal,D.FImporF29,D.NumDocRef"
'    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCtaBanco AND CUE.Codigo = CS.Codigo "
'    Q1 = Q1 & " WHERE CS.Ano = YEAR(D.FEmisionOri - 2) AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCtaBanco"
'    Q1 = Q1 & " ,D.TipoRelEnt,D.IdSucursal,0 as TotPagadoAnoAnt,D.FImportSuc,D.Giro,D.FacCompraRetParcial"
'    Q1 = Q1 & " ,D.IVAIrrecuperable,D.DocOtrosEnAnalitico,D.OldIdDocTmp,D.NumFiscImpr,D.NumInformeZ"
'    Q1 = Q1 & " ,D.CantBoletas,D.VentasAcumInfZ,D.IdDocAsoc,D.PropIVA,D.ValIVAIrrec,D.IVAInmueble"
'    Q1 = Q1 & " ,D.FImpFacturacion,D.CodSIIDTEIVAIrrec,D.TipoDocAsoc,D.IVAActFijo,D.EntRelacionada"
'    Q1 = Q1 & " ,D.NumCuotas,D.CompraBienRaiz,D.NumDocAsoc,D.DTEDocAsoc,D.IdANegCCosto,D.UrlDTE"
'    Q1 = Q1 & " ,D.CodCtaAfectoOld,D.CodCtaExentoOld,D.CodCtaTotalOld,D.DocOtroEsCargo,D.ValRet3Porc"
'    Q1 = Q1 & " ,D.IdCuentaRet3Porc,D.Tratamiento"
'    Q1 = Q1 & " FROM Documento D"
'    Q1 = Q1 & " LEFT JOIN Documento DO ON DO.IdEmpresa = D.IdEmpresa AND DO.Ano =" & gEmpresa.Ano & " AND DO.NumDoc = D.NumDoc AND DO.TipoLib = D.TipoLib AND DO.TipoDoc = D.TipoDoc"
'    Q1 = Q1 & " WHERE D.Ano = " & gEmpresa.Ano + 1
'    Q1 = Q1 & " AND D.IdEmpresa = " & gEmpresa.id
'    Q1 = Q1 & " AND D.FEmisionOri IS NOT NULL AND D.OldIdDoc IS not NULL AND YEAR(D.FEmisionOri - 2) = " & gEmpresa.Ano
'    Q1 = Q1 & " AND DO.IdDoc IS NULL"
'    Q1 = Q1 & " SET IDENTITY_INSERT Documento OFF"
'
'    Rc = ExecSQL(DbMain, Q1)
'
'    If Rc > 0 Then
'      MsgBox1 "Documentos pagados recuperados correctamente.", vbInformation + vbOKOnly
'    End If
'
'    End If
'    'Fin insert documentos
'
'    'MsgBox1 "Documentos pagados recuperados correctamente.", vbInformation + vbOKOnly
'
'
'   Me.MousePointer = vbDefault
'
'End Sub

Private Sub GetDocumentosEliminados() ' esta funcion recupera los documentos eliminados pagados en año anterior
   Dim Rs As Recordset
   Dim Rs2 As Recordset
   Dim Q1 As String
   Dim Q2 As String
   Dim Rc As Long
   
   Me.MousePointer = vbHourglass
    
   
    Q1 = ""
    Q1 = Q1 & " INSERT INTO dbo.MovDocumento"
    Q1 = Q1 & " (IdEmpresa,Ano,IdDoc,IdCompCent,IdCompPago,Orden,IdCuenta,Debe,Haber,Glosa,"
    Q1 = Q1 & " IdTipoValLib,EsTotalDoc,IdCCosto,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld)"
    Q1 = Q1 & " SELECT MV.IdEmpresa," & gEmpresa.Ano & " Ano,DOC.OldIdDoc"
    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV2 WHERE MV2.IdDoc = DOC.OldIdDoc AND MV2.DeCentraliz = 1 AND MV2.Ano = " & gEmpresa.Ano & " AND MV2.IdEmpresa = DOC.IdEmpresa),0) IdCompCent"
    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV1 WHERE MV1.IdDoc = DOC.OldIdDoc AND MV1.DePago = 1 AND MV1.Ano = " & gEmpresa.Ano & " AND MV1.IdEmpresa = DOC.IdEmpresa Order by MV1.IdMov desc),0) IdCompPago"
    Q1 = Q1 & " ,Orden"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = MV.IdCuenta AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = DOC.IdEmpresa),0) IdCuenta"
    Q1 = Q1 & " ,Debe,Haber,Glosa,IdTipoValLib,EsTotalDoc,IdCCosto"
    Q1 = Q1 & " ,IdAreaNeg,Tasa,EsRecuperable,CodSIIDTE,CodCuentaOld"
    Q1 = Q1 & " FROM MovDocumento MV INNER JOIN"
    Q1 = Q1 & " (SELECT D.IdDoc, D.FEmisionOri,D.OldIdDoc, D.IdEmpresa"
    Q1 = Q1 & " FROM Documento D"
    Q1 = Q1 & " LEFT JOIN Documento DO ON DO.IdEmpresa = D.IdEmpresa AND DO.Ano = " & gEmpresa.Ano
    Q1 = Q1 & " AND DO.NumDoc = D.NumDoc AND DO.TipoLib = D.TipoLib AND DO.TipoDoc = D.TipoDoc"
    Q1 = Q1 & " WHERE D.Ano = " & gEmpresa.Ano + 1
    Q1 = Q1 & " AND D.IdEmpresa = " & gEmpresa.id
    Q1 = Q1 & " AND D.FEmisionOri IS NOT NULL AND D.OldIdDoc IS not NULL AND YEAR(D.FEmisionOri -2) <= " & gEmpresa.Ano
    Q1 = Q1 & " AND DO.IdDoc IS NULL) DOC ON DOC.IdDoc = MV.IdDoc"
    
    Rc = ExecSQL(DbMain, Q1)
    'Fin insert mov documento
    
    If Rc > 0 Then
     
     Rc = 0
    '........................
    Q1 = ""
    Q1 = " SET IDENTITY_INSERT Documento ON"
    Q1 = Q1 & " INSERT INTO Documento"
    Q1 = Q1 & " (IdDoc,IdEmpresa,Ano,IdCompCent,IdCompPago,TipoLib,TipoDoc,NumDoc,NumDocHasta"
    Q1 = Q1 & " ,IdEntidad,TipoEntidad,RutEntidad,NombreEntidad,FEmision,FVenc,Descrip,Estado,Exento"
    Q1 = Q1 & " ,IdCuentaExento,Afecto,IdCuentaAfecto,IVA,IdCuentaIVA,OtroImp,IdCuentaOtroImp,Total"
    Q1 = Q1 & " ,IdCuentaTotal,IdUsuario,FechaCreacion,FEmisionOri,CorrInterno,SaldoDoc,FExported"
    Q1 = Q1 & " ,OldIdDoc,DTE,PorcentRetencion,TipoRetencion,MovEdited,OtrosVal,FImporF29,NumDocRef"
    Q1 = Q1 & " ,IdCtaBanco,TipoRelEnt,IdSucursal,TotPagadoAnoAnt,FImportSuc,Giro,FacCompraRetParcial"
    Q1 = Q1 & " ,IVAIrrecuperable,DocOtrosEnAnalitico,OldIdDocTmp,NumFiscImpr,NumInformeZ,CantBoletas"
    Q1 = Q1 & " ,VentasAcumInfZ,IdDocAsoc,PropIVA,ValIVAIrrec,IVAInmueble,FImpFacturacion,CodSIIDTEIVAIrrec"
    Q1 = Q1 & " ,TipoDocAsoc,IVAActFijo,EntRelacionada,NumCuotas,CompraBienRaiz,NumDocAsoc,DTEDocAsoc"
    Q1 = Q1 & " ,IdANegCCosto,UrlDTE,CodCtaAfectoOld,CodCtaExentoOld,CodCtaTotalOld,DocOtroEsCargo"
    Q1 = Q1 & " ,ValRet3Porc,IdCuentaRet3Porc,Tratamiento)"
    Q1 = Q1 & " SELECT D.OldIdDoc IdDoc,D.IdEmpresa, " & gEmpresa.Ano & " AS Ano"
    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV WHERE MV.IdDoc = D.OldIdDoc AND MV.DeCentraliz = 1 AND MV.Ano = YEAR(D.FEmisionOri -2) AND MV.IdEmpresa = D.IdEmpresa),0) IdCompCent"
    Q1 = Q1 & " ,ISNULL((select top 1 IdComp from MovComprobante MV1 WHERE MV1.IdDoc = D.OldIdDoc AND MV1.DePago = 1 AND MV1.Ano = YEAR(D.FEmisionOri -2) AND MV1.IdEmpresa = D.IdEmpresa order by MV1.IdMov desc),0) IdCompPago"
    Q1 = Q1 & " ,D.TipoLib,D.TipoDoc,D.NumDoc,D.NumDocHasta,D.IdEntidad,D.TipoEntidad"
    Q1 = Q1 & " ,D.RutEntidad,D.NombreEntidad,D.FEmision,D.FVenc,D.Descrip,3 Estado,D.Exento"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaExento AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaExento"
    Q1 = Q1 & " ,D.Afecto"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaAfecto AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaAfecto"
    Q1 = Q1 & " ,D.IVA"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaIVA AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaIVA"
    Q1 = Q1 & " ,D.OtroImp"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaOtroImp AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaOtroImp"
    Q1 = Q1 & " ,D.Total"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCuentaTotal AND CUE.Codigo = CS.Codigo"
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCuentaTotal"
    Q1 = Q1 & " ,D.IdUsuario,D.FechaCreacion,D.FEmisionOri,D.CorrInterno,D.SaldoDoc,D.FExported,0 as OldIdDoc"
    Q1 = Q1 & " ,D.DTE,D.PorcentRetencion,D.TipoRetencion,D.MovEdited,D.OtrosVal,D.FImporF29,D.NumDocRef"
    Q1 = Q1 & " ,ISNULL((SELECT CS.idCuenta FROM Cuentas CS INNER JOIN Cuentas CUE ON CUE.IdCuenta = D.IdCtaBanco AND CUE.Codigo = CS.Codigo "
    Q1 = Q1 & " WHERE CS.Ano = " & gEmpresa.Ano & " AND CS.IdEmpresa = D.IdEmpresa),0) AS IdCtaBanco"
    Q1 = Q1 & " ,D.TipoRelEnt,D.IdSucursal,0 as TotPagadoAnoAnt,D.FImportSuc,D.Giro,D.FacCompraRetParcial"
    Q1 = Q1 & " ,D.IVAIrrecuperable,D.DocOtrosEnAnalitico,D.OldIdDocTmp,D.NumFiscImpr,D.NumInformeZ"
    Q1 = Q1 & " ,D.CantBoletas,D.VentasAcumInfZ,D.IdDocAsoc,D.PropIVA,D.ValIVAIrrec,D.IVAInmueble"
    Q1 = Q1 & " ,D.FImpFacturacion,D.CodSIIDTEIVAIrrec,D.TipoDocAsoc,D.IVAActFijo,D.EntRelacionada"
    Q1 = Q1 & " ,D.NumCuotas,D.CompraBienRaiz,D.NumDocAsoc,D.DTEDocAsoc,D.IdANegCCosto,D.UrlDTE"
    Q1 = Q1 & " ,D.CodCtaAfectoOld,D.CodCtaExentoOld,D.CodCtaTotalOld,D.DocOtroEsCargo,D.ValRet3Porc"
    Q1 = Q1 & " ,D.IdCuentaRet3Porc,D.Tratamiento"
    Q1 = Q1 & " FROM Documento D"
    Q1 = Q1 & " LEFT JOIN Documento DO ON DO.IdEmpresa = D.IdEmpresa AND DO.Ano =" & gEmpresa.Ano & " AND DO.NumDoc = D.NumDoc AND DO.TipoLib = D.TipoLib AND DO.TipoDoc = D.TipoDoc"
    Q1 = Q1 & " WHERE D.Ano = " & gEmpresa.Ano + 1
    Q1 = Q1 & " AND D.IdEmpresa = " & gEmpresa.id
    Q1 = Q1 & " AND D.FEmisionOri IS NOT NULL AND D.OldIdDoc IS not NULL AND YEAR(D.FEmisionOri - 2) <= " & gEmpresa.Ano
    Q1 = Q1 & " AND DO.IdDoc IS NULL"
    Q1 = Q1 & " SET IDENTITY_INSERT Documento OFF"

    Rc = ExecSQL(DbMain, Q1)
    
        If Rc > 0 Then
          MsgBox1 "Documentos pagados recuperados correctamente.", vbInformation + vbOKOnly
        End If
    
    
    Else
     MsgBox1 "No existen Documentos a recuperar.", vbInformation + vbOKOnly
    End If
    'Fin insert documentos
    
    'MsgBox1 "Documentos pagados recuperados correctamente.", vbInformation + vbOKOnly


   Me.MousePointer = vbDefault
    
End Sub
'3405779

#End If

Private Sub Form_Terminate()

End Sub

Private Sub GetBotonRecuSQL()
Dim Q1 As String
Dim Rs As Recordset

   Q1 = "SELECT Codigo FROM PARAM "
   Q1 = Q1 & " WHERE Tipo = 'BTRECUSQL' "
   
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
      
    If vFld(Rs("Codigo")) = 1 Then
     Bt_RecuDocuSql.visible = True
    Else
     Bt_RecuDocuSql.visible = False
    End If
   
   End If

End Sub

