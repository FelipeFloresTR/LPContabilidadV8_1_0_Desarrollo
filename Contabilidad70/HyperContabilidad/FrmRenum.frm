VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRenum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renumerar Comprobantes"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "FrmRenum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   1620
      TabIndex        =   9
      Top             =   360
      Width           =   3675
      Begin VB.TextBox tx_nCorrelativo 
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   3
         Top             =   1860
         Width           =   1395
      End
      Begin VB.TextBox tx_nCorrelativo 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   2
         Top             =   1500
         Width           =   1395
      End
      Begin VB.TextBox tx_nCorrelativo 
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   1
         Top             =   1140
         Width           =   1395
      End
      Begin VB.TextBox tx_nCorrelativo 
         Height          =   315
         Index           =   4
         Left            =   1380
         TabIndex        =   0
         Top             =   780
         Width           =   1395
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   2580
         Picture         =   "FrmRenum.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   230
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Apertura:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Traspasos:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Egresos:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Ingresos:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
   End
   Begin MSComctlLib.ProgressBar Pb_Reg 
      Height          =   270
      Left            =   1620
      TabIndex        =   8
      Top             =   2940
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Bt_Close 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Renum 
      Caption         =   "Renumerar"
      Height          =   315
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   420
      Picture         =   "FrmRenum.frx":0316
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "IMPORTANTE: ningún otro usuario debe estar utilizando el sistema mientras se ejecuta esta operación."
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1620
      TabIndex        =   15
      Top             =   3420
      Width           =   5235
   End
End
Attribute VB_Name = "FrmRenum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INC_APERTURA = 1
Const SIN_APERTURA = 1

Private lMinFecha As Long
Dim lTipoCompMod As String

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Desde)
   
   Set Frm = Nothing

End Sub

Private Sub Bt_Renum_Click()
   Dim F1 As Long
   Dim FIni As Long
   Dim ConNewCorr As Boolean
   
   F1 = GetTxDate(Tx_Desde)
   FIni = DateSerial(gEmpresa.Ano, 1, 1)
   
   If F1 <= 0 Then
       MsgBox "Fecha inválida.", vbExclamation
       Exit Sub
   End If
   
   If MsgBox1("Verifique que ningún otro usuario esté trabajando en el sistema." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   If (vFmt(tx_nCorrelativo(TC_EGRESO)) <> 0 Or vFmt(tx_nCorrelativo(TC_INGRESO)) <> 0 Or vFmt(tx_nCorrelativo(TC_APERTURA)) <> 0 Or vFmt(tx_nCorrelativo(TC_TRASPASO)) <> 0) Then
      If F1 <> FIni Then
         MsgBox1 "Si desea renumerar con un correlativo definido por usted, la fecha debe ser 1 Enero " & gEmpresa.Ano, vbExclamation
         Tx_Desde.SetFocus
         Exit Sub
      ElseIf F1 < lMinFecha Then
         If MsgBox1("Ya se imprimieron comprobantes o libros hasta " & FmtFecha(lMinFecha) & "." & vbNewLine & vbNewLine & "¿Desea renumerar de todas maneras?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
   
      ConNewCorr = True
   ElseIf F1 < lMinFecha Then
      If MsgBox1("Solo debería renumerar desde " & FmtFecha(lMinFecha) & " porque ya se imprimieron comprobantes o libros hasta esta fecha." & vbNewLine & vbNewLine & "¿Desea renumerar de todas maneras?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Call SetTxDate(Tx_Desde, lMinFecha)
         Exit Sub
      End If
   End If
      
   MousePointer = vbHourglass
   Bt_Renum.Enabled = False
   DoEvents
   
   Call RenumCompr(F1, ConNewCorr, TAJUSTE_FINANCIERO)
   Call RenumCompr(F1, ConNewCorr, TAJUSTE_TRIBUTARIO)
   
   Bt_Renum.Enabled = True
   MousePointer = vbDefault
   
   MsgBox1 "Proceso terminado.", vbInformation
   
End Sub

Private Sub Form_Load()
   
   Call LoadAll
   
   Call EnabHab(gPerCorrComp = TCC_MENSUAL Or gTipoCorrComp = TCC_UNICO)  'No necesita poner númeración
   Call SetupPriv

End Sub

Private Sub tx_Desde_Change()
   Pb_Reg.Value = 0
End Sub

Private Sub Tx_Desde_GotFocus()
   Call DtGotFocus(Tx_Desde)
End Sub

Private Sub Tx_Desde_LostFocus()
   
   If Trim(Tx_Desde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Desde)
   
End Sub

Private Function RenumCompr(ByVal FMes1 As Long, ConNewCorr As Boolean, ByVal TipoAjuste As Integer)
   Dim Q1 As String, Rs As Recordset, Rc As Long, LastMes As Long, Mes As Long, NCorr As Long
   Dim i As Integer, Tipo As Integer, LastMes1 As Long, LastTipo As Integer
   Dim Corr(TC_APERTURA) As Long, WhTipoComp As String
   Dim WhTAjuste As String
   Dim TmpTbl As String

   
   Pb_Reg.Value = 0
   DoEvents
   WhTipoComp = ""
   
   If TipoAjuste = TAJUSTE_FINANCIERO Then
      WhTAjuste = " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Else
      WhTAjuste = " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
   End If
   
   If ConNewCorr Then
      For i = TC_EGRESO To TC_APERTURA
         If vFmt(tx_nCorrelativo(i)) <> 0 Then 'SACO LOS TIPOS DE COMP. QUE TIENEN Nº PARA MODIFICAR SU CORRELATIVO
            Corr(i) = Val(tx_nCorrelativo(i)) - 1 'Porque despues se le suma uno, por eso se resta
            WhTipoComp = WhTipoComp & "," & i
         End If
      Next i
      
      WhTipoComp = Mid(WhTipoComp, 2)
      WhTipoComp = " AND Tipo IN(" & WhTipoComp & ")"
      
      If gTipoCorrComp = TCC_UNICO Then
         Corr(0) = Corr(TC_APERTURA)
         WhTipoComp = " AND Tipo IN(" & TC_APERTURA & "," & TC_EGRESO & "," & TC_INGRESO & "," & TC_TRASPASO & ")"
      End If

   Else 'ESTO ES COMO ESTABA ANTES SIN LOS NUMEROS SOLO LA FECHA
      For i = 0 To UBound(Corr)
         Corr(i) = 0
      Next i
      
      LastMes1 = -1
         
       'AQUI SE BUSCA DE DONDE PARTE LOS NUMEROS DE COMPROBANTES
      If gPerCorrComp = TCC_ANUAL Or Day(FMes1) <> 1 Then
         Mes = DateSerial(Year(FMes1), month(FMes1), 1)
   
         If gTipoCorrComp = TCC_UNICO Then
            Q1 = "SELECT Max(Correlativo) as M FROM Comprobante WHERE Fecha <" & FMes1 & WhTAjuste
            If gPerCorrComp = TCC_MENSUAL Then
               Q1 = Q1 & " AND Fecha >=" & Mes
               LastMes1 = Year(Mes) * 100 + month(Mes)
            End If
            
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            
            Set Rs = OpenRs(DbMain, Q1)
            If Rs.EOF = False Then
               Corr(0) = vFld(Rs("M"))
            Else
               Corr(0) = 0
            End If
            Call CloseRs(Rs)
            
         Else ' Tipo Comprobante
            Q1 = "SELECT Tipo, Max(Correlativo) as M FROM Comprobante WHERE Fecha <" & FMes1 & WhTAjuste
            If gPerCorrComp = TCC_MENSUAL Then
               Q1 = Q1 & " AND Fecha >=" & Mes
               LastMes1 = Year(Mes) * 100 + month(Mes)
            End If
            
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Q1 = Q1 & " GROUP BY Tipo"
            
            Set Rs = OpenRs(DbMain, Q1)
            Do Until Rs.EOF
               Corr(vFld(Rs("Tipo"))) = vFld(Rs("M"))
               Rs.MoveNext
            Loop
            Call CloseRs(Rs)
               
         End If
         
      End If
   
   End If
   
   Pb_Reg.Value = 1
   
   '******** SE CREA TABLA DE PASO
   TmpTbl = DbGenTmpName2(gDbType, "TRenum_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

   Q1 = "SELECT Tipo, idComp, Correlativo, Fecha, IdEmpresa, Ano INTO " & TmpTbl & " FROM Comprobante WHERE Fecha >=" & FMes1 & WhTAjuste
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & WhTipoComp
   Rc = ExecSQL(DbMain, Q1)
   
   '*******************
   
   '***BARRA PROGRESIVA
   Pb_Reg.Value = Pb_Reg.Value + 1
   
   Q1 = "SELECT Count(*) as N FROM " & TmpTbl
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Pb_Reg.Max = Pb_Reg.Value + vFld(Rs("N"))
   End If
   Call CloseRs(Rs)
   
   '********
   
   Q1 = "SELECT Tipo, Fecha, idComp, Correlativo, IdEmpresa, Ano FROM " & TmpTbl
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   If gTipoCorrComp = TCC_TIPOCOMP Then
      Q1 = Q1 & " ORDER BY Tipo, Fecha, idComp"
   Else
      Q1 = Q1 & " ORDER BY Fecha, idComp"
   End If
      
   LastTipo = -1

   Set Rs = OpenRs(DbMain, Q1)
   Do Until Rs.EOF
   
      If gTipoCorrComp = TCC_TIPOCOMP Then
         Tipo = vFld(Rs("Tipo"))
         If LastTipo <> Tipo Then
            LastTipo = Tipo
            LastMes = LastMes1  ' Volvemos al mes de inicio
         End If
      End If
   
      If gPerCorrComp = TCC_MENSUAL Then
         Mes = vFld(Rs("Fecha"))
         Mes = Year(Mes) * 100 + month(Mes)
         
         If LastMes <> Mes Then 'Si cambia de mes el correlativo vuelve a 1
            LastMes = Mes
            
            If gTipoCorrComp = TCC_TIPOCOMP Then
               Corr(Tipo) = 0
            Else
               Corr(0) = 0
            End If
                        
         End If
      End If
      
      If gTipoCorrComp = TCC_TIPOCOMP Then
         Tipo = vFld(Rs("Tipo"))
         Corr(Tipo) = Corr(Tipo) + 1
         NCorr = Corr(Tipo)
      Else
         Corr(0) = Corr(0) + 1
         NCorr = Corr(0)
      End If
      
      Q1 = "UPDATE Comprobante SET Correlativo=" & NCorr
      Q1 = Q1 & " WHERE idComp=" & vFld(Rs("idComp")) & " AND Correlativo<>" & NCorr
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = ExecSQL(DbMain, Q1)
      
      Dim Where As String
      Where = " WHERE idComp=" & vFld(Rs("idComp")) & " AND Correlativo<>" & NCorr & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call SeguimientoComprobantes(0, gEmpresa.id, gEmpresa.Ano, "FrmConfigCorrComp.Bt_MarcarRes_Click", "", 1, Where, gUsuario.IdUsuario, 1, 2)

   
      Pb_Reg.Value = Pb_Reg.Value + 1
   
      Debug.Print Q1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

End Function

Private Sub SetupPriv()

   If Not ChkPriv(PRV_ADM_COMP) Then
      Bt_Renum.Enabled = False
   End If
End Sub
Private Sub LoadAll()
   Dim Q1 As String, Rs As Recordset
      
   Q1 = "SELECT Max(FHasta) as Fecha FROM LogImpreso WHERE Estado <> " & EL_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   lMinFecha = vFld(Rs("Fecha"))
   Call CloseRs(Rs)
   
   If lMinFecha <= 0 Then
      lMinFecha = DateSerial(gEmpresa.Ano, 1, 1)
   Else
      lMinFecha = lMinFecha + 1
   End If

   Call SetTxDate(Tx_Desde, lMinFecha)
   
End Sub
'Desahabilita botones, porque no necesita ingresar nº de comprobantes
Private Sub EnabHab(bool As Boolean)

    Call SetTxRO(tx_nCorrelativo(TC_EGRESO), bool)
    Call SetTxRO(tx_nCorrelativo(TC_INGRESO), bool)
    Call SetTxRO(tx_nCorrelativo(TC_TRASPASO), bool)
    
    If gPerCorrComp = TCC_ANUAL And gTipoCorrComp = TCC_UNICO Then
         bool = False
    End If
    Call SetTxRO(tx_nCorrelativo(TC_APERTURA), bool)
    
End Sub
