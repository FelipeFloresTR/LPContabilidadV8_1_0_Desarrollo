VERSION 5.00
Begin VB.Form FrmImportRemu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar desde Remuneraciones"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "FrmImportRemu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Generar"
      Height          =   315
      Left            =   6360
      TabIndex        =   5
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   540
      Picture         =   "FrmImportRemu.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Generar Comprobante Remuneraciones "
      Height          =   3315
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   4575
      Begin VB.CheckBox Ch_DesglozarCCosto 
         Caption         =   "Identificar Centros de Costo/Gestión"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2100
         Width           =   3975
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1020
         Width           =   3075
      End
      Begin VB.CheckBox Ch_ResAnual 
         Caption         =   "Generar Resumen Anual"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2580
         Width           =   3975
      End
      Begin VB.CheckBox Ch_DetEmpl 
         Caption         =   "Remuneraciones por pagar Detallado por Empleado"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C. Costo:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   9
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   480
         Width           =   345
      End
   End
End
Attribute VB_Name = "FrmImportRemu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SG_PASSW_FAIRPAY = "oP,*/'#2j7h7_$3"


Dim lEsLPRemu As Boolean
Dim lRemuSQLServer As Boolean

Dim lMsgAdv As Boolean
Dim lCtasRemu(MAX_CTASREMU) As Long
Dim lIdEmpresaRem As Long
Dim lPathlDbRemu As String
Dim lConnStr As String
Dim lEmpSep As Boolean
Dim lCbCCosto As ClsCombo
Dim lDesglozarCCosto As Boolean

Const COD_CCOSTO = 2

'Remu Ley 21.227
Const MOVP_SUSPAUTOR = 13             ' Suspensión por acto de autoridad, Ley 21227 (23 abr 2020)
Const MOVP_SUSPPACT = 14              ' Suspensión por pacto, Ley 21227 (23 abr 2020)
Const MOVP_REDJORN = 15               ' Reducción de jornada laboral, Ley 21227 (23 abr 2020)


Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim i As Integer
      
   If lMsgAdv = False Then    'este mensaje se muestra sólo una vez
   
      If MsgBox1("Para realizar la importación desde Remuneraciones, nadie debe estar trabajando en esta empresa en Remuneraciones." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   
   End If
   
   lMsgAdv = True
      
   Me.MousePointer = vbHourglass
   
   Call GenCompRemu(ItemData(Cb_Mes), gEmpresa.Ano)
         
   Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim DbF29Path As String
   
   lMsgAdv = False
   
   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   
   Tx_Ano = gEmpresa.Ano
   
   If OpenlDbRemu() = False Then
      MsgBox1 "Problemas al abrir la base de datos de Remuneraciones.", vbExclamation
      Bt_Import.Enabled = False
      Exit Sub
   End If
   
   
   Set lCbCCosto = New ClsCombo
   Call lCbCCosto.SetControl(Cb_CCosto)
   
   Call lCbCCosto.AddItem("(todos)", "-1", "-1")
   Q1 = "SELECT CCosto & ' - ' & Descrip as Item, idCCosto, CCosto FROM CCostos WHERE idEmpresa=" & lIdEmpresaRem
   Q1 = Q1 & " AND CCosto <> ' ' AND NOT CCosto IS NULL ORDER BY CCosto"
'   Call lCbCCosto.FillCombo(lDbRemu, Q1, "-1")
   Set Rs = OpenRsDao(lDbRemu, Q1)
   Do While Not Rs.EOF
      Call lCbCCosto.AddItem(vFldDao(Rs("Item")), vFldDao(Rs("IdCCosto")), vFldDao(Rs("CCosto")))
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   lCbCCosto.ListIndex = -1
      
   lDesglozarCCosto = False
      
End Sub
Private Function GenCompRemu(ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal idcomp As Long = 0) As Long
   Dim Q1 As String, Q2 As String
   Dim Rs As dao.Recordset, RsAux As Recordset
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer, nLiq As Integer
   Dim IdCompNew As Long
   Dim NomMes As String
   Dim Fecha As Long
   Dim MesActual As Integer
   Dim AnoMes As Long
   Dim QBase As String
   Dim QEnd As String
   Dim RemuPagar As Double
   Dim Tot As Double
   Dim Tipo As Integer
   Dim Glosa As String
   Dim Idx As Integer, VersionDbRem As Integer
   Dim WhAnoMes As String
   Dim RemAnual As Boolean
   Dim TxtPeriodo As String
   Dim AuxPathlDbRemu As String
   Dim IdCCostoRemu As Long
   Dim WhCCosto As String
   Dim ValFonasa As Double
   Const P_VERSIONBD = "'VERSIONBD'"
   Dim FirstDay As Long, LastDay As Long
   Dim QEndCCosto As String
   Dim IdCCostoContab As Long, CCostoContab As String
   Dim SumBonoNoImp As Long, SumBonoImp As Long
   Dim TmpTblComp As String
   Dim FldArray(3) As AdvTbAddNew_t
   
   GenCompRemu = 0
      
   MesActual = GetMesActual()   'debe haber un mes abierto
   
   If MesActual = 0 Then
      MsgBox "No hay mes abierto.", vbExclamation
      Exit Function
   End If
      
'   Q1 = "SELECT Codigo FROM Rem_Param WHERE Tipo=" & P_VERSIONBD
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'      VersionDbRem = vflddao(Rs("Codigo"))
'
'   End If
'   Call CloseRs(Rs)
'
'   If VersionDbRem < 55 Then
'      MsgBox "La versión de Remuneraciones debe ser 3.0.14 o superior para poder generar el comprobante de remuneraciones.", vbExclamation
'      Exit Function
'   End If
      
      
   AnoMes = gEmpresa.Ano * 100# + CbItemData(Cb_Mes)
   
   If Ch_DesglozarCCosto <> 0 Then
      lDesglozarCCosto = True
   Else
      lDesglozarCCosto = False
   End If

   
   If Ch_ResAnual <> 0 Then
      RemAnual = True
      WhAnoMes = " WHERE (r.AnoMes BETWEEN " & gEmpresa.Ano * 100# + 1 & " AND " & gEmpresa.Ano * 100# + 12 & ")"
   Else
      RemAnual = False
      WhAnoMes = " WHERE r.AnoMes = " & AnoMes
   End If
   
   WhCCosto = ""
   IdCCostoRemu = lCbCCosto.ItemData
   If IdCCostoRemu > 0 Then
      WhCCosto = " AND r.idCCosto = " & IdCCostoRemu
   Else
      IdCCostoRemu = 0
   End If
   
   IdCCostoContab = 0
   CCostoContab = ""
   If IdCCostoRemu > 0 Then
      Q1 = "SELECT IdCCosto, Descripcion FROM CentroCosto WHERE Codigo = '" & lCbCCosto.Matrix(COD_CCOSTO, lCbCCosto.ListIndex) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Set RsAux = OpenRs(DbMain, Q1)
      If Not RsAux.EOF Then
         IdCCostoContab = vFld(RsAux("IdCCosto"))
         CCostoContab = vFld(RsAux("Descripcion"))
      Else
         MsgBox1 "Centro de gestión no encontrado en Contabilidad." & vbCrLf & vbCrLf & "El código del Centro de Costo de Remuneraciones debe coincidir exactamente con el código del Centro de Gestión en Contabilidad.", vbExclamation
         Call CloseRs(RsAux)
         Exit Function
      End If
      Call CloseRs(RsAux)
   End If

      
   If GetEstadoMes(CbItemData(Cb_Mes)) <> EM_ABIERTO Then
      MsgBox1 "El mes seleccionado no está abierto.", vbExclamation
      Exit Function
   End If
   
   
   Call LoadCuentas
   
   If lCtasRemu(REMU_CTASUELDOBASE) = 0 Or lCtasRemu(REMU_CTAIMPUNICO) = 0 Or lCtasRemu(REMU_CTAREMPAGAR) = 0 Then
      MsgBox1 "Falta configurar las cuentas para el traspaso de Remuneraciones." & vbCrLf & vbCrLf & "Utilice la opción: " & vbCrLf & vbCrLf & "        'Configuración Traspaso Remuneraciones'" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If
     
   Q1 = "SELECT Count(*) FROM Resumen as r "
   Q1 = Q1 & WhAnoMes
   Q1 = Q1 & " AND IdEmpresa = " & lIdEmpresaRem
   Q1 = Q1 & WhCCosto
   
   Set Rs = OpenRsDao(lDbRemu, Q1)
   
   If Rs.EOF = False Then
      If vFldDao(Rs(0)) = 0 Then
         MsgBox1 "No es posible generar el comprobante de Remuneraciones debido a que no hay liquidaciones emitidas para el periodo y centro de costo seleccionados.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   
   Call CloseRs(Rs)
   
   If RemAnual Then
      TxtPeriodo = " año " & Tx_Ano
   Else
      TxtPeriodo = " mes de " & Cb_Mes & " " & Tx_Ano
   End If
   
   If MsgBox1("Se iniciará el proceso de generación del comprobante contable correspondiente al " & TxtPeriodo & "." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
      Exit Function
   End If
                        
   Me.MousePointer = vbHourglass
   
   'creamos el comprobante
   If idcomp = 0 Then

      FldArray(0).FldName = "IdUsuario"
      FldArray(0).FldValue = gUsuario.IdUsuario
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "FechaCreacion"
      FldArray(1).FldValue = CLng(Int(Now))
      FldArray(1).FldIsNum = True
            
      FldArray(2).FldName = "IdEmpresa"
      FldArray(2).FldValue = gEmpresa.id
      FldArray(2).FldIsNum = True
                  
      FldArray(3).FldName = "Ano"
      FldArray(3).FldValue = gEmpresa.Ano
      FldArray(3).FldIsNum = True
      
      IdCompNew = AdvTbAddNewMult(DbMain, "Comprobante", "IdComp", FldArray)
   End If
   
      
   'OJO: Cuando se cambia esta función se debe cambiar la función GenDetRemuPagar y el Form del FairPay: FrmAsientoContable
   
   ' 7 abr 2014. por pequeños cambios
   Q1 = "SELECT 1 as x "
   If IdCCostoRemu <= 0 And lDesglozarCCosto Then
      Q1 = Q1 & ", r.idCCosto "
   End If
   Q1 = Q1 & ", Sum(r.SueldoBase) as SumSueldoBase "
   Q1 = Q1 & ", Sum(r.HorasExt) + Sum(r.HorasExt100) + Sum(r.OtrasHrsExt) As SumHorasExtra "
   Q1 = Q1 & ", Sum(r.BonoDomingo) + Sum(r.BonoDomingoEx) As SumBonoDomingo "
   Q1 = Q1 & ", Sum(r.Comision) as SumComision "
'   Q1 = Q1 & ", Sum(r.Bono) as SumBonoImp" ' 8 jun 2015: se agrega BonoDomingo
'   Q1 = Q1 & ", Sum(r.BonoNoImp) as SumBonoNoImp"

   ' 3 feb 2017: se usa la vista QBono
   Q1 = Q1 & ", Sum(b.BImpTrib) + Sum(b.BImpNoTrib) as SumBonoImp"
   Q1 = Q1 & ", Sum(b.BNoImpTrib) + Sum(b.BNoImpNoTrib) as SumBonoNoImp "

   Q1 = Q1 & ", Sum(Gratif) as SumGratif "
   Q1 = Q1 & ", Sum(r.ValSemCorr) as SumValSemCorr "     'Semana Corrida / 4 jun 2012
   Q1 = Q1 & ", Sum(r.Moviliz) as SumMoviliz "
   Q1 = Q1 & ", Sum(r.Colac) as SumColacion "
   Q1 = Q1 & ", Sum(r.AsigFamiliar) as SumAsigFam "
   Q1 = Q1 & ", Sum(r.CargasRetro) as SumCargasRetro "

'   Q1 = Q1 & ", Sum(Isapre + AdicIsapre) as SumIsapre "
'   Q1 = Q1 & ", Sum(r.Isapre) as SumIsapre "   ' pam: 24 nov 2010: ya incluye el adicional
   Q1 = Q1 & ", Sum(iif(isp.IdIsapre <> 1, r.Isapre + r.CotSalud21227, 0)) as SumIsapre "
'   Q1 = Q1 & ", Sum(" & SqlCase(gDbType, "isp.IdIsapre <> 1", "r.Isapre+r.CotSalud21227", "0") & ") as SumIsapre"
'   Q1 = Q1 & ", Sum(iif(isp.IdIsapre = 1, r.Isapre - r.CCAF, 0)) as SumFonasa "   ' pam: 4 jun 2012
   Q1 = Q1 & ", Sum(iif(isp.IdIsapre = 1 OR isp.IdIsapre Is NULL, r.Isapre + r.CotSalud21227, 0)) as SumFonasa "   ' pam: 13 dic 2013: se agrega IS NULL por si no exiete el registro en IsapreEmpl
'   Q1 = Q1 & ", Sum( " & SqlCase(gDbType, "isp.IdIsapre Is NULL OR isp.IdIsapre = 1", "r.Isapre+r.CotSalud21227", "0") & ") as SumFonasa"

'   Q1 = Q1 & ", Sum(iif(h.idAFP > 0, AFP - CotSisEmpl, 0) + AFPCuenta2 + AfVolCap + AfVolAhorro) as SumAFP "
   Q1 = Q1 & ", Sum(iif(h.idAFP > 0, r.AFP - r.CotSisEmpl, 0) + r.AFPCuenta2 + r.AfVolCap + r.AfVolAhorro) + Sum(r.CotAfp21227) as SumAFP"

   Q1 = Q1 & ", Sum(iif(h.idAFP > 0, 0, AFP)) as SumINP "
'   Q1 = Q1 & ", Sum(SegAcc) as SumSegAcc "
   Q1 = Q1 & ", Sum(r.SegAcc) + Sum(r.CotAcc21227) as SumSegAcc"  ' 6 may 2020: se agrega 21227
   Q1 = Q1 & ", Sum(r.SegAccEmpl) as SumSegAccEmpl"  ' 26 may 2014: para Socios

'   Q1 = Q1 & ", Sum(SegCes) as SumSegCesEmpl "
'   Q1 = Q1 & ", Sum(SegCesEmpl) as SumSegCesEmpr "
   Q1 = Q1 & ", Sum(r.SegCes) + Sum(r.CotSCesTrab21227) as SumSegCesEmpl"
   Q1 = Q1 & ", Sum(r.SegCesEmpl) + Sum(r.CotSCesEmpr21227) as SumSegCesEmpr"

   Q1 = Q1 & ", Sum(CotSisEmpl) as SumSISEmpl "
'   Q1 = Q1 & ", Sum(CotSisEmpr) as SumSISEmpr "
   Q1 = Q1 & ", Sum(r.CotSisEmpr) + Sum(r.CotSIS21227) as SumSISEmpr"

   Q1 = Q1 & ", Sum(APV) as SumAPV"
   Q1 = Q1 & ", Sum(APVCEmpl) as SumAPVCEmpl"
   Q1 = Q1 & ", Sum(APVCEmpr) as SumAPVCEmpr"
'   Q1 = Q1 & ", Sum(CCAF) as SumCCAF"
   Q1 = Q1 & ", Sum(r.CCAF) + Sum(CotCCAF21227) as SumCCAF"

   Q1 = Q1 & ", Sum(CredCCAFValor) as SumCredCCAF "
   Q1 = Q1 & ", Sum(OtrosDesc) as SumDescuentos "
   Q1 = Q1 & ", Sum(r.ValNoTrab)  as SumValNoTrab "        'SumNoTrabajado = SumValNoTrab + SumDescAtraso
   Q1 = Q1 & ", Sum(r.DescAtraso) as SumDescAtraso " ' Se calcula por separado porque con los NULL falla la suma
   Q1 = Q1 & ", Sum(s.MayorRet) as SumMayorRet "
   Q1 = Q1 & ", Sum(Impto) as SumImpto "
   Q1 = Q1 & ", Sum(Anticipo) as SumAnticipo "
   Q1 = Q1 & ", Sum(Prestamo) as SumPrestamo "
   Q1 = Q1 & ", Sum(CotTPesado) as SumCotTPesado "
   
   '2758884 se agrega filtro que sea mayor al año 2021
   If gEmpresa.Ano >= 2021 Then
    Q1 = Q1 & ", Sum(r.Retenc3P) as SumRetenc3P "
   End If
   ' fin 2758884
   
   Q1 = Q1 & ", Sum(iif(SubLiquido < 0, -SubLiquido, 0)) as SumSubLiquido "

   ' 26 may 2020: Cotizaciones que paga el empleador por ley 21227
   Q1 = Q1 & ", Sum(iif(qm.TipoMov =" & MOVP_REDJORN & ",qm.ImponPactado,0)) as ImpPact21227"
   Q1 = Q1 & ", Sum(iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),iif(isp.IdIsapre <> 1,r.CotSalud21227,0),0))"
   Q1 = Q1 & "+ Sum(iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),iif(isp.IdIsapre Is NULL OR isp.IdIsapre = 1,r.CotSalud21227,0),0))"
   Q1 = Q1 & "+ Sum(iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.CotAfp21227,0))"
   Q1 = Q1 & "+ Sum(iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.CotSCesTrab21227,0)) as Cotiz21227"
   Q1 = Q1 & ", Sum(iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.AsigFamiliar,0)) as AF21227"
   
   'pipe 2791437
   Q1 = Q1 & ", Sum(MontoCovid) as SEGCOVID "


   Q1 = Q1 & " FROM ((((Resumen as r INNER JOIN Sueldos as s ON r.AnoMes = s.AnoMes AND r.idEmpl = s.idEmpl AND r.idEmpresa = s.idEmpresa)"
   Q1 = Q1 & " INNER JOIN EmplHist h ON r.AnoMes = h.AnoMes AND r.idEmpl = h.idEmpl AND r.idEmpresa = h.idEmpresa)" ' 31 jul 2017: pam: se agrega
   Q1 = Q1 & " LEFT JOIN QBonos as b ON r.AnoMes = b.AnoMes AND r.idEmpl = b.idEmpl And r.idEmpresa = b.idEmpresa)"
   Q1 = Q1 & " LEFT JOIN IsapreEmp as isp ON r.AnoMes = isp.AnoMes AND r.idEmpl = isp.idEmpl And r.idEmpresa = isp.idEmpresa)"
   Q1 = Q1 & " LEFT JOIN QMov21227 qm ON r.AnoMes = qm.AnoMes AND r.idEmpl = qm.idEmpl And r.idEmpresa = qm.idEmpresa"
   Q1 = Q1 & WhAnoMes
   Q1 = Q1 & " AND r.IdEmpresa = " & lIdEmpresaRem
   Q1 = Q1 & WhCCosto

   If IdCCostoRemu <= 0 And lDesglozarCCosto Then
      Q1 = Q1 & " GROUP BY r.idCCosto "
      Q1 = Q1 & " ORDER BY r.idCCosto "
   End If

         
   Set Rs = OpenRsDao(lDbRemu, Q1)
   
   'primero creamos la tabla vacía
   TmpTblComp = DbGenTmpName2(SQL_ACCESS, "tmpremu_MovComp_")   'forzamos Access para que no use # en el nombre temporal
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblComp)
   
   Q1 = "SELECT MovComprobante.*, 0 as IdTipoRemu INTO " & TmpTblComp & " FROM MovComprobante WHERE 1=0 "
   Call ExecSQL(DbMain, Q1, False)
   
   QBase = "INSERT INTO " & TmpTblComp
   QBase = QBase & "(IdComp, IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, IdCartola, DeCentraliz, DePago, DeRemu, IdTipoRemu) "
   QBase = QBase & " VALUES(" & IdCompNew & ", 0, "
   
   
   
   QEnd = ",0,0,0,0,0,1,"
   QEndCCosto = "," & IdCCostoContab & ",0,0,0,0,1,"
   Glosa = ""
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   i = 1
   RemuPagar = 0
   nLiq = 0
      
   Do While Not Rs.EOF
      nLiq = nLiq + 1
      
      If IdCCostoRemu <= 0 And lDesglozarCCosto Then
         
         IdCCostoContab = 0
         CCostoContab = ""
         
         If vFldDao(Rs("idCCosto")) > 0 Then
            
            Idx = lCbCCosto.FindItem(vFldDao(Rs("idCCosto")))
            Q2 = "SELECT IdCCosto, Descripcion FROM CentroCosto WHERE Codigo = '" & lCbCCosto.Matrix(COD_CCOSTO, Idx) & "'"
            Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id
            Set RsAux = OpenRs(DbMain, Q2)
            
            If Not RsAux.EOF Then
               IdCCostoContab = vFld(RsAux("IdCCosto"))
               CCostoContab = vFld(RsAux("Descripcion"))
               QEndCCosto = "," & IdCCostoContab & ",0,0,0,0,1,"

            Else
               MsgBox1 "Centro de gestión no encontrado en Contabilidad." & vbCrLf & vbCrLf & "El código del Centro de Costo de Remuneraciones debe coincidir exactamente con el código del Centro de Gestión en Contabilidad.", vbExclamation
               Call CloseRs(RsAux)
               Call CloseRs(Rs)
               
               Call ExecSQL(DbMain, "DROP TABLE " & TmpTblComp)
               If gDbType = SQL_SERVER Then
                  Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")    'para que vuelva a mostrar Warnings en el caso de tener que truncar un texto
               End If
   
               Exit Function
            
            End If
            
            Call CloseRs(RsAux)
            
            WhCCosto = " AND r.idCCosto = " & vFldDao(Rs("idCCosto"))
            
         Else   'para los que no están en ningún centro de costo
            WhCCosto = " AND ( r.idCCosto <= 0 OR r.idCCosto IS NULL)"
            
         End If
      End If
      
      RemuPagar = 0
            
      If vFldDao(Rs("SumSueldoBase")) <> 0 And lCtasRemu(REMU_CTASUELDOBASE) <> 0 Then
      
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASUELDOBASE) & "," & vFldDao(Rs("SumSueldoBase")) & ",0,'" & gTipoDatosRemu(REMU_CTASUELDOBASE) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASUELDOBASE, IdCCostoContab)
            
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumSueldoBase"))
      End If
      
      If vFldDao(Rs("ImpPact21227")) <> 0 And lCtasRemu(REMU_CTAIMPPACT) <> 0 Then
      
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAIMPPACT) & "," & vFldDao(Rs("ImpPact21227")) & ",0,'" & gTipoDatosRemu(REMU_CTAIMPPACT) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAIMPPACT, IdCCostoContab)
            
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("ImpPact21227"))
      End If
      
      If vFldDao(Rs("SumHorasExtra")) <> 0 And lCtasRemu(REMU_CTAHORASEXTRA) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAHORASEXTRA) & "," & vFldDao(Rs("SumHorasExtra")) & ",0,'" & gTipoDatosRemu(REMU_CTAHORASEXTRA) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAHORASEXTRA, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumHorasExtra"))
      End If
      
      If vFldDao(Rs("SumBonoDomingo")) <> 0 And lCtasRemu(REMU_CTAHORASDOMINGO) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAHORASDOMINGO) & "," & vFldDao(Rs("SumBonoDomingo")) & ",0,'" & gTipoDatosRemu(REMU_CTAHORASDOMINGO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAHORASDOMINGO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumBonoDomingo"))
      End If
      
      
      If vFldDao(Rs("SumComision")) <> 0 And lCtasRemu(REMU_CTACOMISION) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACOMISION) & "," & vFldDao(Rs("SumComision")) & ",0,'" & gTipoDatosRemu(REMU_CTACOMISION) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACOMISION, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumComision"))
      End If

      SumBonoImp = Round(vFldDao(Rs("SumBonoImp")), 0)
      If SumBonoImp <> 0 And lCtasRemu(REMU_CTABONOSIMP) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTABONOSIMP) & "," & SumBonoImp & ",0,'" & gTipoDatosRemu(REMU_CTABONOSIMP) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTABONOSIMP, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + SumBonoImp
      End If
      
      SumBonoNoImp = Round(vFldDao(Rs("SumBonoNoImp")), 0)
      If SumBonoNoImp <> 0 And lCtasRemu(REMU_CTABONOSNOIMP) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTABONOSNOIMP) & "," & SumBonoNoImp & ",0,'" & gTipoDatosRemu(REMU_CTABONOSNOIMP) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTABONOSNOIMP, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + SumBonoNoImp
      End If
      
      If vFldDao(Rs("SumGratif")) <> 0 And lCtasRemu(REMU_CTAGRATIFIC) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAGRATIFIC) & "," & vFldDao(Rs("SumGratif")) & ",0,'" & gTipoDatosRemu(REMU_CTAGRATIFIC) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAGRATIFIC, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumGratif"))
      End If
      
      If vFldDao(Rs("SumValSemCorr")) <> 0 And lCtasRemu(REMU_CTASEMANACORRIDA) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEMANACORRIDA) & "," & vFldDao(Rs("SumValSemCorr")) & ",0,'" & gTipoDatosRemu(REMU_CTASEMANACORRIDA) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEMANACORRIDA, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumValSemCorr"))
      End If
      
      If vFldDao(Rs("SumMoviliz")) <> 0 And lCtasRemu(REMU_CTAMOVILIZ) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAMOVILIZ) & "," & vFldDao(Rs("SumMoviliz")) & ",0,'" & gTipoDatosRemu(REMU_CTAMOVILIZ) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAMOVILIZ, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("SumMoviliz"))
      End If
      
      If vFldDao(Rs("SumColacion")) <> 0 And lCtasRemu(REMU_CTACOLACION) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACOLACION) & "," & vFldDao(Rs("SumColacion")) & ",0,'" & gTipoDatosRemu(REMU_CTACOLACION) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACOLACION, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumColacion"))
      End If
      
      If vFldDao(Rs("SumAsigFam")) <> 0 And lCtasRemu(REMU_CTAASIGFAM) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAASIGFAM) & "," & vFldDao(Rs("SumAsigFam")) & ",0,'" & gTipoDatosRemu(REMU_CTAASIGFAM) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAASIGFAM, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumAsigFam"))
      End If
      
      If vFldDao(Rs("SumCargasRetro")) <> 0 And lCtasRemu(REMU_CTACARGASRETRO) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACARGASRETRO) & "," & vFldDao(Rs("SumCargasRetro")) & ",0,'" & gTipoDatosRemu(REMU_CTACARGASRETRO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACARGASRETRO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumCargasRetro"))
      End If
      
      If vFldDao(Rs("SumSegAcc")) <> 0 And lCtasRemu(REMU_CTASEGACCGASTO) <> 0 Then              'Seg Acc gasto (se repite en el haber)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEGACCGASTO) & "," & vFldDao(Rs("SumSegAcc")) & ",0,'" & gTipoDatosRemu(REMU_CTASEGACCGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEGACCGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("SumSegAcc"))
      End If

      
      If vFldDao(Rs("SumSegCesEmpr")) <> 0 And lCtasRemu(REMU_CTASEGCESEMPRGASTO) <> 0 Then          'Seg Cesantía Empresa (se repite en el haber)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEGCESEMPRGASTO) & "," & vFldDao(Rs("SumSegCesEmpr")) & ",0,'" & gTipoDatosRemu(REMU_CTASEGCESEMPRGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEGCESEMPRGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("SumSegCesEmpr"))
      End If
      
      If vFldDao(Rs("SumSISEmpr")) <> 0 And lCtasRemu(REMU_CTASISEMPRGASTO) <> 0 Then        'SIS Empresa (se repite en el haber)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASISEMPRGASTO) & "," & vFldDao(Rs("SumSISEmpr")) & ",0,'" & gTipoDatosRemu(REMU_CTASISEMPRGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASISEMPRGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumSISEmpr"))
      End If
      
      
      'pipe 2791437
      If vFldDao(Rs("SEGCOVID")) <> 0 And lCtasRemu(REMU_CTASSEGCOVIDGASTO) <> 0 Then        'SEG COVID (se repite en el haber)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASSEGCOVIDGASTO) & "," & vFldDao(Rs("SEGCOVID")) & ",0,'" & gTipoDatosRemu(REMU_CTASSEGCOVIDGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASSEGCOVIDGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SEGCOVID"))
      End If
       ' FIN pipe 2791437
      
      If vFldDao(Rs("SumCotTPesado")) <> 0 And lCtasRemu(REMU_CTATRPESADOGASTO) <> 0 Then        'Trabajo pesado (gasto)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTATRPESADOGASTO) & "," & vFldDao(Rs("SumCotTPesado")) & ",0,'" & gTipoDatosRemu(REMU_CTATRPESADOGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTATRPESADOGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar + vFldDao(Rs("SumCotTPesado"))
      End If
            
'      If vflddao(Rs("SumCCAF")) <> 0 And lCtasRemu(REMU_CTACCAF_GASTO) <> 0 Then    'se elimina por petición de Alejandro Contreras (4 jun 2012)
'         Q1 = QBase & i & ","
'         Q1 = Q1 & lCtasRemu(REMU_CTACCAF_GASTO) & "," & vflddao(Rs("SumCCAF")) & ",0,'" & gTipoDatosRemu(REMU_CTACCAF_GASTO) & "'"
'
'         Q1 = Q1 & AddCCosto(REMU_CTACCAF_GASTO, IdCCostoContab)
'         Call ExecSQL(DbMain, Q1, false)
'         i = i + 1
'
'         RemuPagar = RemuPagar + vflddao(Rs("SumCCAF"))
'      End If


      
      If vFldDao(Rs("SumAPVCEmpr")) <> 0 And lCtasRemu(REMU_CTAAPVCEMPRGASTO) <> 0 Then           'APVC gasto (se repite en el haber)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAAPVCEMPRGASTO) & "," & vFldDao(Rs("SumAPVCEmpr")) & ",0,'" & gTipoDatosRemu(REMU_CTAAPVCEMPRGASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAAPVCEMPRGASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("SumAPVCEmpr"))
      End If

      If vFldDao(Rs("Cotiz21227")) <> 0 And lCtasRemu(REMU_CTACOTIZ21227GASTO) <> 0 Then           'Cotizaciones Ley 21227 para las supensiones
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACOTIZ21227GASTO) & "," & vFldDao(Rs("Cotiz21227")) & ",0,'" & gTipoDatosRemu(REMU_CTACOTIZ21227GASTO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACOTIZ21227GASTO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("Cotiz21227"))
      End If

      'Diferencias por cobrar
      
      If vFldDao(Rs("SumSubLiquido")) <> 0 And lCtasRemu(REMU_CTADIFPORCOBRAR) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTADIFPORCOBRAR) & "," & vFldDao(Rs("SumSubLiquido")) & ",0,'" & gTipoDatosRemu(REMU_CTADIFPORCOBRAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTADIFPORCOBRAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar + vFldDao(Rs("SumSubLiquido"))
      End If

      '***********************   aquí terminan los valores al DEBE *************************
      
      
      Tot = RemuPagar
      
      If vFldDao(Rs("SumAFP")) <> 0 And lCtasRemu(REMU_CTAAFP) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAAFP) & ", 0," & vFldDao(Rs("SumAFP")) & ",'" & gTipoDatosRemu(REMU_CTAAFP) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAAFP, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumAFP"))
      End If
      
      If vFldDao(Rs("SumFonasa")) <> 0 And lCtasRemu(REMU_CTAFONASA) <> 0 Then
         ValFonasa = vFldDao(Rs("SumFonasa")) - vFldDao(Rs("SumCCAF"))
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAFONASA) & ",0," & ValFonasa & ",'" & gTipoDatosRemu(REMU_CTAFONASA) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAFONASA, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - ValFonasa
      
      End If
      
      If vFldDao(Rs("SumIsapre")) <> 0 And lCtasRemu(REMU_CTAISAPRE) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAISAPRE) & ",0," & vFldDao(Rs("SumIsapre")) & ",'" & gTipoDatosRemu(REMU_CTAISAPRE) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAISAPRE, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumIsapre"))
      
      End If
      
      If vFldDao(Rs("SumINP")) <> 0 And lCtasRemu(REMU_CTAINP) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAINP) & ",0," & vFldDao(Rs("SumINP")) & ",'" & gTipoDatosRemu(REMU_CTAINP) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAINP, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumINP"))
      
      End If
      
      If vFldDao(Rs("SumSegAcc")) <> 0 And lCtasRemu(REMU_CTASEGACCPPAGAR) <> 0 Then           'Seg Acc. por pagar (se repite en el Debe)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEGACCPPAGAR) & ",0," & vFldDao(Rs("SumSegAcc")) & ",'" & gTipoDatosRemu(REMU_CTASEGACCPPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEGACCPPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumSegAcc"))
      End If

      If vFldDao(Rs("SumSegCesEmpl")) <> 0 And lCtasRemu(REMU_CTASEGCESEMPL) <> 0 Then        'Seg Cesantía Empleado
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEGCESEMPL) & ",0," & vFldDao(Rs("SumSegCesEmpl")) & ",'" & gTipoDatosRemu(REMU_CTASEGCESEMPL) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEGCESEMPL, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumSegCesEmpl"))
      End If

      If vFldDao(Rs("SumSegCesEmpr")) <> 0 And lCtasRemu(REMU_CTASEGCESEMPRPPAGAR) <> 0 Then          'Seg Cesantía Empresa (se repite en el debe)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASEGCESEMPRPPAGAR) & ",0," & vFldDao(Rs("SumSegCesEmpr")) & ",'" & gTipoDatosRemu(REMU_CTASEGCESEMPRPPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASEGCESEMPRPPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumSegCesEmpr"))
      End If
      
      If vFldDao(Rs("SumSISEmpl")) <> 0 And lCtasRemu(REMU_CTASISEMPL) <> 0 Then        'SIS Empleado
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASISEMPL) & ",0," & vFldDao(Rs("SumSISEmpl")) & ",'" & gTipoDatosRemu(REMU_CTASISEMPL) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASISEMPL, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumSISEmpl"))
      End If

'      If vFldDao(Rs("SumCotTPesado")) <> 0 And lCtasRemu(REMU_CTATRPESADOPORPAGAR) <> 0 Then        'Trabajo pesado a pagar
'         Q1 = QBase & i & ","
'         Q1 = Q1 & lCtasRemu(REMU_CTATRPESADOPORPAGAR) & ",0," & vFldDao(Rs("SumCotTPesado")) & ",'" & gTipoDatosRemu(REMU_CTATRPESADOPORPAGAR) & "'"
'
'         Q1 = Q1 & AddCCosto(REMU_CTATRPESADOPORPAGAR, IdCCostoContab)
'
'         Call ExecSQL(DbMain, Q1, False)
'         i = i + 1
'
'         RemuPagar = RemuPagar - vFldDao(Rs("SumCotTPesado"))
'      End If
      
      If vFldDao(Rs("SumCotTPesado")) <> 0 And lCtasRemu(REMU_CTATRPESADOEMPRESA) <> 0 Then        'Trabajo pesado Empresa a pagar
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTATRPESADOEMPRESA) & ",0," & vFldDao(Rs("SumCotTPesado")) & ",'" & gTipoDatosRemu(REMU_CTATRPESADOEMPRESA) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTATRPESADOEMPRESA, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumCotTPesado"))
      End If
      
      If vFldDao(Rs("SumCotTPesado")) <> 0 And lCtasRemu(REMU_CTATRPESADOTRABAJADOR) <> 0 Then        'Trabajo pesado Trabajador a pagar
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTATRPESADOTRABAJADOR) & ",0," & vFldDao(Rs("SumCotTPesado")) & ",'" & gTipoDatosRemu(REMU_CTATRPESADOTRABAJADOR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTATRPESADOTRABAJADOR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumCotTPesado"))
      End If

      If vFldDao(Rs("SumSISEmpr")) <> 0 And lCtasRemu(REMU_CTASISEMPRPPAGAR) <> 0 Then        'SIS Empresa (se repite en el debe)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASISEMPRPPAGAR) & ",0," & vFldDao(Rs("SumSISEmpr")) & ",'" & gTipoDatosRemu(REMU_CTASISEMPRPPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASISEMPRPPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumSISEmpr"))
      End If
      
      
       'pipe 2791437
       
       If vFldDao(Rs("SEGCOVID")) <> 0 And lCtasRemu(REMU_CTASSEGCOVIDPORPAGAR) <> 0 Then        'SEG COVID (se repite en el DEBE)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTASSEGCOVIDPORPAGAR) & ",0," & vFldDao(Rs("SEGCOVID")) & ",'" & gTipoDatosRemu(REMU_CTASSEGCOVIDPORPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTASSEGCOVIDPORPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SEGCOVID"))
      End If
         'FIN pipe 2791437
                

      If vFldDao(Rs("SumAPV")) <> 0 And lCtasRemu(REMU_CTAAPV) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAAPV) & ",0," & vFldDao(Rs("SumAPV")) & ",'" & gTipoDatosRemu(REMU_CTAAPV) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAAPV, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumAPV"))
      End If

      If vFldDao(Rs("SumAPVCEmpl")) <> 0 And lCtasRemu(REMU_CTAAPVCEMPL) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAAPVCEMPL) & ",0," & vFldDao(Rs("SumAPVCEmpl")) & ",'" & gTipoDatosRemu(REMU_CTAAPVCEMPL) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAAPVCEMPL, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumAPVCEmpl"))
      End If

      If vFldDao(Rs("SumAPVCEmpr")) <> 0 And lCtasRemu(REMU_CTAAPVCEMPRPPAGAR) <> 0 Then            'APVC por pagar (se repite en el debe)
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAAPVCEMPRPPAGAR) & ",0," & vFldDao(Rs("SumAPVCEmpr")) & ",'" & gTipoDatosRemu(REMU_CTAAPVCEMPRPPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAAPVCEMPRPPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumAPVCEmpr"))
      End If

      If vFldDao(Rs("SumCCAF")) <> 0 And lCtasRemu(REMU_CTACCAF_PORPAGAR) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACCAF_PORPAGAR) & ",0," & vFldDao(Rs("SumCCAF")) & ",'" & gTipoDatosRemu(REMU_CTACCAF_PORPAGAR) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACCAF_PORPAGAR, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumCCAF"))
      End If

      If vFldDao(Rs("SumCredCCAF")) <> 0 And lCtasRemu(REMU_CTACREDCCAF) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTACREDCCAF) & ",0," & vFldDao(Rs("SumCredCCAF")) & ",'" & gTipoDatosRemu(REMU_CTACREDCCAF) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTACREDCCAF, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumCredCCAF"))
      End If

      If vFldDao(Rs("SumDescuentos")) <> 0 And lCtasRemu(REMU_OTROSDESCUENTOS) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_OTROSDESCUENTOS) & ",0," & vFldDao(Rs("SumDescuentos")) & ",'" & gTipoDatosRemu(REMU_OTROSDESCUENTOS) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_OTROSDESCUENTOS, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumDescuentos"))
      End If

      If (vFldDao(Rs("SumValNoTrab")) + vFldDao(Rs("SumDescAtraso"))) <> 0 And lCtasRemu(REMU_CTAVALNOTRABAJADO) <> 0 Then
         Q1 = QBase & i & ","
         'Q1 = Q1 & lCtasRemu(REMU_CTAVALNOTRABAJADO) & ",0," & (vFldDao(Rs("SumValNoTrab")) + vFldDao(Rs("SumDescAtraso"))) & ",'" & gTipoDatosRemu(REMU_CTAVALNOTRABAJADO) & "'"
         Q1 = Q1 & lCtasRemu(REMU_CTAVALNOTRABAJADO) & ",0," & (vFldDao(Rs("SumValNoTrab")) + vFldDao(Rs("SumDescAtraso"))) & ",'" & Mid(gTipoDatosRemu(REMU_CTAVALNOTRABAJADO), 1, 50) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAVALNOTRABAJADO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - (vFldDao(Rs("SumValNoTrab")) + vFldDao(Rs("SumDescAtraso")))
      End If

      If vFldDao(Rs("SumMayorRet")) <> 0 And lCtasRemu(REMU_CTAMAYORRETENCION) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAMAYORRETENCION) & ",0," & vFldDao(Rs("SumMayorRet")) & ",'" & gTipoDatosRemu(REMU_CTAMAYORRETENCION) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAMAYORRETENCION, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumMayorRet"))
      End If
      
      If vFldDao(Rs("SumImpto")) <> 0 And lCtasRemu(REMU_CTAIMPUNICO) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAIMPUNICO) & ",0," & vFldDao(Rs("SumImpto")) & ",'" & gTipoDatosRemu(REMU_CTAIMPUNICO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAIMPUNICO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
      
         RemuPagar = RemuPagar - vFldDao(Rs("SumImpto"))
      End If
      
      ' 2758884 Se agrega filtro año sea superior al año 2021
      If gEmpresa.Ano >= 2021 Then
        If vFldDao(Rs("SumRetenc3P")) <> 0 And lCtasRemu(REMU_CTARET3PORC) <> 0 Then
           Q1 = QBase & i & ","
           Q1 = Q1 & lCtasRemu(REMU_CTARET3PORC) & ",0," & vFldDao(Rs("SumRetenc3P")) & ",'" & gTipoDatosRemu(REMU_CTARET3PORC) & "'"
           
           Q1 = Q1 & AddCCosto(REMU_CTARET3PORC, IdCCostoContab)
           
           Call ExecSQL(DbMain, Q1, False)
           i = i + 1
        
           RemuPagar = RemuPagar - vFldDao(Rs("SumRetenc3P"))
        End If
      End If
      
      If vFldDao(Rs("SumAnticipo")) <> 0 And lCtasRemu(REMU_CTAANTICIPO) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAANTICIPO) & ", 0," & vFldDao(Rs("SumAnticipo")) & ",'" & gTipoDatosRemu(REMU_CTAANTICIPO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAANTICIPO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumAnticipo"))
      End If

      If vFldDao(Rs("SumPrestamo")) <> 0 And lCtasRemu(REMU_CTAPRESTAMO) <> 0 Then
         Q1 = QBase & i & ","
         Q1 = Q1 & lCtasRemu(REMU_CTAPRESTAMO) & ", 0," & vFldDao(Rs("SumPrestamo")) & ",'" & gTipoDatosRemu(REMU_CTAPRESTAMO) & "'"
         
         Q1 = Q1 & AddCCosto(REMU_CTAPRESTAMO, IdCCostoContab)
         
         Call ExecSQL(DbMain, Q1, False)
         i = i + 1
         
         RemuPagar = RemuPagar - vFldDao(Rs("SumPrestamo"))
      End If
      
      If (RemuPagar <> 0 Or vFldDao(Rs("AF21227")) <> 0) And lCtasRemu(REMU_CTAREMPAGAR) <> 0 Then
      
         If Ch_DetEmpl = 0 Then
            Q1 = QBase & i & ","
            Q1 = Q1 & lCtasRemu(REMU_CTAREMPAGAR) & ",0," & RemuPagar & ",'" & gTipoDatosRemu(REMU_CTAREMPAGAR) & "'"
            
            Q1 = Q1 & AddCCosto(REMU_CTAREMPAGAR, IdCCostoContab)
         
            Call ExecSQL(DbMain, Q1, False)
            i = i + 1
            
         Else

            Call GenDetRemuPagar(WhAnoMes, WhCCosto, QBase, IdCCostoContab, i)
            
         End If
            
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
'   If Ch_DetEmpl <> 0 Then
'
'      Call GenDetRemuPagar(WhAnoMes, WhCCosto, QBase, IdCCostoContab, i)
'
'   End If
   
   Call AddLog("GenCompRemu: Periodo: " & TxtPeriodo & ", nLiq=" & nLiq & ", nMov=" & i - 1)
   
   'ahora insertamos ordenado por tipo de Cuenta Remu en MovComprobante
   Q1 = "INSERT INTO MovComprobante "
   Q1 = Q1 & "( IdComp, IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, IdCartola, DeCentraliz, DePago, DeRemu, IdEmpresa, Ano )"
   Q1 = Q1 & " SELECT IdComp, IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, IdCartola, DeCentraliz, DePago, DeRemu, " & gEmpresa.id & "," & gEmpresa.Ano
   Q1 = Q1 & " FROM " & TmpTblComp
   Q1 = Q1 & " ORDER BY IdTipoRemu"

   Call ExecSQL(DbMain, Q1, False)
   
   Q1 = "UPDATE MovComprobante SET Orden = 0 WHERE IdComp= " & IdCompNew
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1, False)
   
   'actualizamos el encabezado
'   Fecha = CLng(Int(DateSerial(gEmpresa.Ano, CbItemData(Cb_Mes), 15)))
   If MesActual = month(Now) Then
      Fecha = CLng(Int(DateSerial(gEmpresa.Ano, month(Now), Day(Now))))
   Else
      Fecha = DateSerial(gEmpresa.Ano, MesActual, 1)
   End If

   If gDtCompCent <> DTCOMPCENT_CURRDEF Then
      If gDtCompCent = DTCOMPCENT_DEFDAY And (month(Fecha) <> 2 And gDayDtCompCent > 0 And gDayDtCompCent <= 30) Or (month(Fecha) = 2 And gDayDtCompCent > 0 And gDayDtCompCent <= 28) Then
         Fecha = DateSerial(Year(Fecha), month(Fecha), gDayDtCompCent)
      Else
         Call FirstLastMonthDay(Fecha, FirstDay, LastDay)
         Fecha = DateSerial(Year(Fecha), month(Fecha), Day(LastDay))
      End If
   End If
   
   If Mes >= 1 And Mes <= 12 Then
      NomMes = gNomMes(Mes)
   End If
   
   Tipo = TC_TRASPASO
   
   If RemAnual Then
      Glosa = "Centraliz. Remuneraciones Año " & Ano
   Else
      Glosa = "Centraliz. Remuneraciones Mes de " & NomMes & " " & Ano
   End If
   
   If IdCCostoContab > 0 And IdCCostoRemu > 0 Then
      Glosa = Glosa & " - C. Costo: " & CCostoContab
   End If
      
   
   Q1 = "UPDATE Comprobante SET "
   Q1 = Q1 & "  Fecha = " & Fecha
   Q1 = Q1 & ", Tipo = " & Tipo
   Q1 = Q1 & ", TipoAjuste = " & TAJUSTE_AMBOS
   Q1 = Q1 & ", Estado = " & gEstadoNewComp
   Q1 = Q1 & ", Glosa = '" & Glosa & "'"
   Q1 = Q1 & ", TotalDebe = " & Tot
   Q1 = Q1 & ", TotalHaber = " & Tot
   Q1 = Q1 & ", ImpResumido = " & Abs(gImpResCent)
   Q1 = Q1 & " WHERE IdComp = " & IdCompNew
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Call ExecSQL(DbMain, Q1, False)
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblComp)

   idcomp = IdCompNew
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")
   End If
   
         
   Me.MousePointer = vbDefault
   
   MsgBox1 "ATENCIÓN: si le falta algún desglose de cuentas en el comprobante, verifique que la cuenta esté definida en la opción de menú " & vbCrLf & vbCrLf & "Configuración >> Configuración para Traspaso a Remuneraciones", vbInformation

   If idcomp > 0 Then
      If RemAnual Then
         Call FrmComprobante.FEditCentraliz(idcomp, 0, Ano)
      Else
         Call FrmComprobante.FEditCentraliz(idcomp, Mes, Ano)
      End If
   Else
      MsgBox1 "Problemas al generar el comprobante.", vbExclamation + vbOKOnly
   End If
   
   GenCompRemu = idcomp
End Function
Private Sub LoadCuentas()
   Dim i As Integer
   Dim Tipo As Integer
   
   For i = 1 To UBound(gTipoDatosRemu)
      If gTipoDatosRemu(i) <> "" Then
         lCtasRemu(i) = Val(GetParamEmpresa("CTASREMU", i))
      End If
   Next i

End Sub

Private Sub GenDetRemuPagar(ByVal WhAnoMes As String, ByVal WhCCosto As String, ByVal QBase As String, ByVal IdCCostoContab As Long, ByVal orden As Integer)
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim RemuPagar As Double
   Dim i As Integer
   Dim QEnd As String, QEndCCosto As String, QEndC As String
   Dim Ano As Integer, Mes As Integer
   
   If lCtasRemu(REMU_CTAREMPAGAR) = 0 Then
      Exit Sub
   End If

   Q1 = "SELECT r.AnoMes, r.idEmpl, r.IdCCosto, " & SqlNombre() & ","
   Q1 = Q1 & " r.SueldoBase as SumSueldoBase, "
   Q1 = Q1 & " r.HorasExt, r.HorasExt100, r.OtrasHrsExt, "
   Q1 = Q1 & " r.BonoDomingo, r.BonoDomingoEx, "
   Q1 = Q1 & " r.Comision as SumComision, "
'   Q1 = Q1 & " r.Bono as SumBonoImp, "
'   Q1 = Q1 & " r.BonoNoImp as SumBonoNoImp, "
   
   ' 3 feb 2017: se usa la vista QBono
   Q1 = Q1 & " (b.BImpTrib + b.BImpNoTrib) as SumBonoImp,"
   Q1 = Q1 & " (b.BNoImpTrib + b.BNoImpNoTrib) as SumBonoNoImp,"

   Q1 = Q1 & " Gratif as SumGratif,"
   Q1 = Q1 & " r.ValSemCorr as SumValSemCorr,"     '4 jun 2012
   Q1 = Q1 & " r.Moviliz as SumMoviliz,"
   Q1 = Q1 & " r.Colac as SumColacion,"
   Q1 = Q1 & " r.AsigFamiliar as SumAsigFam,"
   Q1 = Q1 & " r.CargasRetro as SumCargasRetro,"
   
'   Q1 = Q1 & " r.Isapre + AdicIsapre as SumIsapre, "
'   Q1 = Q1 & " r.Isapre as SumIsapre, " ' pam: 24 nuv 2010: ya incluye el adicional
   Q1 = Q1 & " iif(isp.IdIsapre <> 1, r.Isapre + r.CotSalud21227, 0) as SumIsapre, "   ' pam: 4 jun 2012
'   Q1 = Q1 & " iif(isp.IdIsapre = 1, r.Isapre - CCAF, 0) as SumFonasa, "   ' pam: 4 jun 2012
   Q1 = Q1 & " iif(isp.IdIsapre = 1 or isp.IdIsapre Is NULL, r.Isapre + r.CotSalud21227, 0) as SumFonasa, "   ' pam: 4 jun 2012
    
'   Q1 = Q1 & " iif(s.idAFP > 0, AFP - CotSisEmpl, 0) + AFPCuenta2 + AfVolCap + AfVolAhorro as SumAFP, "
'   Q1 = Q1 & " iif(s.idAFP > 0, 0, AFP) as SumINP, "
   Q1 = Q1 & " iif(h.idAFP > 0, AFP - CotSisEmpl, 0) + AFPCuenta2 + AfVolCap + AfVolAhorro + r.CotAFP21227 as SumAFP, "
   Q1 = Q1 & " iif(h.idAFP > 0, 0, AFP) as SumINP, "
   Q1 = Q1 & " SegAcc + r.CotAcc21227 as SumSegAcc, "
   Q1 = Q1 & " SegCes + r.CotSCesTrab21227 as SumSegCesEmpl, "
   Q1 = Q1 & " SegCesEmpl + r.CotSCesEmpr21227 as SumSegCesEmpr, "
   Q1 = Q1 & " CotSisEmpl as SumSISEmpl, "
   Q1 = Q1 & " CotSisEmpr + r.CotSIS21227 as SumSISEmpr, "
   Q1 = Q1 & " APV as SumAPV,"
   Q1 = Q1 & " APVCEmpl as SumAPVCEmpl,"
   Q1 = Q1 & " APVCEmpr as SumAPVCEmpr,"
   Q1 = Q1 & " CCAF - r.CotCCAF21227 as SumCCAF,"
   Q1 = Q1 & " CredCCAFValor as SumCredCCAF,"
   Q1 = Q1 & " OtrosDesc as SumDescuentos,"
   Q1 = Q1 & " ValNoTrab  as SumValNoTrab,"        'SumNoTrabajado = SumValNoTrab + SumDescAtraso
   Q1 = Q1 & " r.DescAtraso as SumDescAtraso,"
   Q1 = Q1 & " s.MayorRet as SumMayorRet,"
   Q1 = Q1 & " Impto as SumImpto, r.Retenc3P as SumRetenc3P, "
   Q1 = Q1 & " Anticipo as SumAnticipo,"
   Q1 = Q1 & " Prestamo as SumPrestamo, "
   Q1 = Q1 & " CotTPesado as SumTPesado, "
   Q1 = Q1 & " iif( SubLiquido < 0, -SubLiquido, 0 ) as SumSubLiquido "
   
  ' 26 may 2020: Cotizaciones que paga el empleador por ley 21227
   Q1 = Q1 & ", iif(qm.TipoMov =" & MOVP_REDJORN & ",qm.ImponPactado,0) as ImpPact21227"
   Q1 = Q1 & ", iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),iif(isp.IdIsapre <> 1,r.CotSalud21227,0),0)"
   Q1 = Q1 & "+ iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),iif(isp.IdIsapre Is NULL OR isp.IdIsapre = 1,r.CotSalud21227,0),0)"
   Q1 = Q1 & "+ iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.CotAfp21227,0)"
   Q1 = Q1 & "+ iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.CotSCesTrab21227,0) as Cotiz21227"
   Q1 = Q1 & ", iif(qm.TipoMov IN (" & MOVP_SUSPAUTOR & "," & MOVP_SUSPPACT & "),r.AsigFamiliar,0) as AF21227"
  
   Q1 = Q1 & " FROM (((((Resumen as r INNER JOIN Sueldos as s ON r.AnoMes = s.AnoMes AND r.idEmpl = s.idEmpl AND r.idEmpresa = s.idEmpresa)"
   Q1 = Q1 & " INNER JOIN EmplHist h ON r.AnoMes = h.AnoMes AND r.idEmpl = h.idEmpl AND r.idEmpresa = h.idEmpresa)" ' 31 jul 2017: pam: se agrega
   Q1 = Q1 & " INNER JOIN Empleados ON r.idEmpl = Empleados.IdEmpl And r.idEmpresa = Empleados.idEmpresa)"
   Q1 = Q1 & " LEFT JOIN QBonos as b ON r.AnoMes = b.AnoMes AND r.idEmpl = b.idEmpl And r.idEmpresa = b.idEmpresa)"
   Q1 = Q1 & " LEFT JOIN IsapreEmp as isp ON r.AnoMes = isp.AnoMes AND r.idEmpl = isp.idEmpl)"
   Q1 = Q1 & " LEFT JOIN QMov21227 qm ON r.AnoMes = qm.AnoMes AND r.idEmpl = qm.idEmpl And r.idEmpresa = qm.idEmpresa"
   
   
   Q1 = Q1 & WhAnoMes
   Q1 = Q1 & " AND r.IdEmpresa = " & lIdEmpresaRem
   Q1 = Q1 & WhCCosto
   
   Q1 = Q1 & " ORDER BY " & SqlNombre(False) & ", r.AnoMes"
         
   Set Rs = OpenRsDao(lDbRemu, Q1)
   
   i = orden
   
   QEnd = ",0,0,0,0,0,1," & REMU_CTAREMPAGAR & ")"
   QEndCCosto = "," & IdCCostoContab & ",0,0,0,0,1," & REMU_CTAREMPAGAR & ")"

   If IdCCostoContab > 0 Then
      If GetAtribCuenta(lCtasRemu(REMU_CTAREMPAGAR), ATRIB_CCOSTO) <> 0 Then
         QEndC = QEndCCosto
      Else
         QEndC = QEnd
      End If
   Else
      QEndC = QEnd
   End If

   Do While Not Rs.EOF
     
      RemuPagar = 0

      RemuPagar = RemuPagar + vFldDao(Rs("SumSueldoBase"))
      RemuPagar = RemuPagar + vFldDao(Rs("ImpPact21227"))
      RemuPagar = RemuPagar + vFldDao(Rs("HorasExt")) + vFldDao(Rs("HorasExt100")) + vFldDao(Rs("OtrasHrsExt"))
      RemuPagar = RemuPagar + vFldDao(Rs("BonoDomingo")) + vFldDao(Rs("BonoDomingoEx"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumComision"))
      RemuPagar = RemuPagar + Round(vFldDao(Rs("SumBonoImp")), 0)
      RemuPagar = RemuPagar + Round(vFldDao(Rs("SumBonoNoImp")), 0)
      RemuPagar = RemuPagar + vFldDao(Rs("SumGratif"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumValSemCorr"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumMoviliz"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumColacion"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumAsigFam"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumCargasRetro"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumSegAcc"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumSegCesEmpr"))    'se repite en haber
      RemuPagar = RemuPagar + vFldDao(Rs("SumSISEmpr"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumAPVCEmpr"))
      RemuPagar = RemuPagar + vFldDao(Rs("Cotiz21227"))
      RemuPagar = RemuPagar + vFldDao(Rs("SumSubLiquido"))

      RemuPagar = RemuPagar - vFldDao(Rs("SumAFP"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumIsapre"))
      RemuPagar = RemuPagar - (vFldDao(Rs("SumFonasa")) - vFldDao(Rs("SumCCAF")))
      RemuPagar = RemuPagar - vFldDao(Rs("SumINP"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumSegAcc"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumSegCesEmpl"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumSegCesEmpr"))     'se repite en debe
      RemuPagar = RemuPagar - vFldDao(Rs("SumSISEmpl"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumSISEmpr"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumAPV"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumAPVCEmpl"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumAPVCEmpr"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumCCAF"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumCredCCAF"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumDescuentos"))
      RemuPagar = RemuPagar - (vFldDao(Rs("SumValNoTrab")) + vFldDao(Rs("SumDescAtraso")))
      RemuPagar = RemuPagar - vFldDao(Rs("SumMayorRet"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumImpto"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumRetenc3P"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumAnticipo"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumPrestamo"))
      RemuPagar = RemuPagar - vFldDao(Rs("SumTPesado"))

      Q1 = QBase & i & ","
      If Ch_ResAnual <> 0 Then
         Ano = vFldDao(Rs("AnoMes")) \ 100
         Mes = vFldDao(Rs("AnoMes")) Mod 100
         Q1 = Q1 & lCtasRemu(REMU_CTAREMPAGAR) & ",0," & RemuPagar & ",'" & FCase(vFldDao(Rs("Nombre"))) & " (" & Format(Mes, "mm") & "-" & Ano & ")'" & QEndC
      Else
         Q1 = Q1 & lCtasRemu(REMU_CTAREMPAGAR) & ",0," & RemuPagar & ",'" & FCase(vFldDao(Rs("Nombre"))) & "'" & QEndC
      End If
      
      Call ExecSQL(DbMain, Q1)
      i = i + 1
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

End Sub
Private Function AddCCosto(ByVal IdTipoCuenta As Long, ByVal IdCCosto As Long) As String
   Dim QEnd As String
   Dim QEndCCosto As String
   
   QEnd = ",0,0,0,0,0,1," & IdTipoCuenta & ")"
   QEndCCosto = "," & IdCCosto & ",0,0,0,0,1," & IdTipoCuenta & ")"

   If IdCCosto > 0 Then
      If GetAtribCuenta(lCtasRemu(IdTipoCuenta), ATRIB_CCOSTO) <> 0 Then
         AddCCosto = QEndCCosto
      Else
         AddCCosto = QEnd
      End If
   Else
      AddCCosto = QEnd
   End If

End Function

Public Function SqlNombre(Optional bAs As Boolean = True) As String
   Dim Buf As String

'   Buf = "Trim(Empleados.ApPaterno & ' ' & Empleados.ApMaterno) & ', ' & Empleados.Nombres & IIF(Empleados.Codigo = ' ',' ',' (' & Empleados.Codigo & ')' )"
   Buf = "(Empleados.ApPaterno & ' ' & Empleados.ApMaterno) & ', ' & Empleados.Nombres & IIF(Empleados.Codigo = ' ',' ',' (' & Empleados.Codigo & ')' )"

   If bAs Then
      Buf = Buf & " as Nombre"
   End If
   
   SqlNombre = Buf

End Function

Private Function OpenlDbRemu() As Boolean
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim Cfg As String
   Dim AuxPathlDbRemu As String
   Dim Idx As Integer
   Dim Qry As QueryDef
   Dim RazonSoc As String
   Dim Buf As String
   Dim TmpQry As String


   OpenlDbRemu = False

   lPathlDbRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   Buf = GetIniString(gIniFile, "Config", "VersionRemu", "")
   lRemuSQLServer = IIf(UCase(Buf) = "SQLSERVER", True, False)
   
   If lPathlDbRemu = "" Then
      MsgBox1 "Falta configurar la localización del sistema de Remuneraciones. Utilice la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If
   
   If Not ExistFile(lPathlDbRemu) Then
      MsgBox1 "No se encontró el archivo de " & IIf(lRemuSQLServer, "configuración", "base dee datos") & " del sistema de Remuneraciones." & vbCrLf & vbCrLf & "Verifique la localización del archivo en la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If
   
   If lRemuSQLServer Then
      If OpenMsSqlRemu() = False Then
         Exit Function
      End If
      
   Else
   
#If DATACON = 1 Then
      Cfg = GetIniString(gCfgFile, "Config", "Secur", "")
      If Cfg <> SG_SEGCFG Then         'suponenemos que si la base de LPContab no tiene clave, tampoco la tiene la de FairPay
         lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
      Else
         lConnStr = ""
      End If
   
      lConnStr = Mid(lConnStr, 2)
   
#Else
      lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"    'si estamos en SQLServer, se asume que la base de Remu Access tiene  clave

#End If
   
      On Error Resume Next
      
      Set lDbRemu = OpenDatabase(lPathlDbRemu, False, False, lConnStr)
      
      If lDbRemu Is Nothing Then
         Call AddLog("OpenDbRemu: Error " & ERR & ", " & Error & ", " & lPathlDbRemu)
         
         lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"    'asumimos que faltó la clave (caso particular de Fairware)
             
         Set lDbRemu = OpenDatabase(lPathlDbRemu, False, False, lConnStr)
               
         
         If lDbRemu Is Nothing Then
            MsgBox1 "No se pudo abrir la base de datos " & vbCrLf & lPathlDbRemu & vbCrLf & vbCrLf & Error, vbExclamation
            Call AddLog("OpenDbRemu: Error " & ERR & ", " & Error & ", " & lPathlDbRemu)
            Exit Function
         End If
      End If
         
   End If
   
   'en remuneraciones se permite tener dos empresas con el mismo RUT, por eso se da la opción al cliente de seleccionar cuál de las dos (9 abril 2019)
   Q1 = "SELECT IdEmpresa, RazonSoc FROM Empresas WHERE Rut = '" & gEmpresa.Rut & "'"
   'Q1 = "SELECT sum(ValNoTrab) FROM Resumen WHERE AnoMes = 202208 and idEmpresa = 26 "
   Set Rs = OpenRsDao(lDbRemu, Q1)
   If Not Rs.EOF Then
      lIdEmpresaRem = vFldDao(Rs("IdEmpresa"))
      RazonSoc = vFldDao(Rs("RazonSoc"))

      Rs.MoveNext

      If Not Rs.EOF Then
         If MsgBox1("Existen dos empresas con este mismo RUT." & vbCrLf & "¿Desea obtener los datos de '" & RazonSoc & "'?", vbQuestion + vbYesNo) = vbNo Then
            lIdEmpresaRem = vFldDao(Rs("IdEmpresa"))
         End If
      End If
      
   End If
   Call CloseRs(Rs)


   If lIdEmpresaRem = 0 Then
      MsgBox "No se ha encontrado esta empresa en el sistema de Remuneraciones.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If Not lRemuSQLServer Then
   
      Q1 = "SELECT Codigo FROM Param WHERE Tipo = 'EMPSEP'"
      Set Rs = OpenRsDao(lDbRemu, Q1)
      If Rs.EOF = False Then
         lEmpSep = IIf(vFldDao(Rs("Codigo")) <> 0, True, False)
      Else
         lEmpSep = False
      End If
      Call CloseRs(Rs)
      
      lDbRemu.Close
      Set lDbRemu = Nothing
         
      AuxPathlDbRemu = lPathlDbRemu
      
      
      If lEmpSep Then
         Idx = InStrRev(AuxPathlDbRemu, "\")
         If Idx > 0 Then
            AuxPathlDbRemu = Left(AuxPathlDbRemu, Idx)
         End If
    
         AuxPathlDbRemu = AuxPathlDbRemu & "Empresas\" & gEmpresa.Rut & "_" & lIdEmpresaRem & ".mdb"
         'SF 14263961
          If lDbRemu Is Nothing Then
            lConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"    'asumimos que faltó la clave (caso particular de Fairware
          End If
          'SF 14263961
      End If

' se descomenta para quitar la clave a la base de datos.
'        Set lDbRemu = OpenDatabase(AuxPathlDbRemu, True, False, lConnStr)
'        lDbRemu.NewPassword SG_PASSW_FAIRPAY, ""
'        Call CloseDb(lDbRemu)
'        lConnStr = ""
        Set lDbRemu = OpenDatabase(AuxPathlDbRemu, False, False, lConnStr)
      If lDbRemu Is Nothing Then
         Call AddLog("OpenDbRemu : Error " & ERR & ", " & Error & ", " & AuxPathlDbRemu)
         Exit Function
      End If

   End If
   
   OpenlDbRemu = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   
   Call CloseDb(lDbRemu)
   
   Set lDbRemu = Nothing
   

End Sub


' Para MsSql Server
' Verificar SHOW VARIABLES LIKE 'lower_case_table_names'  que sea 1 o 2
Function OpenMsSqlRemu() As Boolean
   Dim Rc As Integer, SqlPort As Long, Usr As String, Psw As String, i As Integer
   Dim ConnStr As String, Host As String, UsrPsw As String, DbName As String
   Dim sErr1 As Long, sError1 As String, Encript As Boolean, CfgFile As String

   On Error Resume Next
   
   OpenMsSqlRemu = False

   If Not lDbRemu Is Nothing Then
      lDbRemu.Close
      Set lDbRemu = Nothing
   End If
   
   CfgFile = lPathlDbRemu
   If LCase(Right(lPathlDbRemu, 10)) = "lpremu.cfg" Then
      lEsLPRemu = True
   ElseIf LCase(Right(lPathlDbRemu, 11)) = "fairpay.cfg" Then
      lEsLPRemu = False
   Else
      MsgBox1 "Falta especificar correctamente el archivo de configuración de Remuneraciones." & vbCrLf & "Utilice la opción " & vbCrLf & vbCrLf & "Configuración Traspaso Remuneraciones" & vbCrLf & vbCrLf & "bajo el menú Configuración.", vbExclamation
      Exit Function
   End If
   
   Host = Trim(GetIniString(CfgFile, "MS Sql", "Host", ""))
   
   If Host = "" Then
      MsgBox1 "Falta especificar el servidor de base de datos." & vbCrLf & "Comuníquese con su administrador.", vbCritical
      Exit Function
   End If
      
   SqlPort = Val(GetIniString(CfgFile, "MS Sql", "Port", "1433"))
      
   If lEsLPRemu Then
      Debug.Print "Db lpremu=" & FwEncrypt1("               lpremu             ", 56516)
      DbName = GetIniString(CfgFile, "MS Sql", "DB", FwDecrypt1("6E2C6B2B6C2E71357A40874F98E2D8D7DFDBDA5E2F8154287D532A825B35906C4927", 56516))

      Usr = GetIniString(CfgFile, "MS Sql", "User", "lp" & "re" & "mu")
   Else
      Debug.Print "Db fairpay=" & FwEncrypt1("           fairpay           ", 56516)
      DbName = GetIniString(CfgFile, "MS Sql", "DB", FwDecrypt1("9053975C2269317A448F5BA89DABABAAB3C553287E552D86603B977452", 56516))
   
      Usr = GetIniString(CfgFile, "MS Sql", "User", "fai" & "rp" & "ay")
   End If
   
   
   Debug.Print "Hola Psw=" & FwEncrypt1("     " & DbName & "   #" & "      hola       ", 731982) ' ojo con el #
   Debug.Print "Oficial Psw=" & FwEncrypt1("     " & DbName & "   #" & "     _F&].[r94%.        ", 731982) ' ojo con el #
   
   Psw = GetIniString(CfgFile, "MS Sql", "Pswk")
   
   If Psw = "" Then
      MsgBox1 "Falta especificar la clave del servidor de base de datos de Remuneraciones." & vbCrLf & "Comuníquese con su administrador.", vbCritical
      Exit Function
   End If
   
   Psw = Trim(FwDecrypt1(Psw, 731982))
   i = InStr(Psw, "#")
   Psw = Trim(Mid(Psw, i + 1))
         
   UsrPsw = "U" & "ID=" & Usr & ";P" & "WD=" & Psw & ";"
      
   ConnStr = "Driver={SQL Server};Server=" & Host & ";MARS_Connection=yes;Database=" & DbName & ";" ' 2 abr 2018
        
   On Error Resume Next

   Set lDbRemu = OpenDatabase("", False, False, ConnStr & UsrPsw)
      
'   Set lDbRemu = New Connection
'   lDbRemu.ConnectionString = ConnStr & UsrPsw
'   lDbRemu.Open
      
   If ERR Then
      If ERR <> 3059 Then
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & ConnStr, vbCritical
      End If
      Call AddLog("OpenMsSqlRemu: Error " & ERR & ", " & Error & ", " & ConnStr)
      
      Set lDbRemu = Nothing
      
      End
      Exit Function
   End If
    
   If ERR Then
      sErr1 = ERR.Number
      sError1 = ERR.Description
      MsgErr "Verifique que esté bien definido el servidor de la base de datos y que tenga los privilegios necesarios."
      Call AddLog("Error " & sErr1 & ", " & sError1 & ", [" & ConnStr & "]")
   Else
      OpenMsSqlRemu = True
      
      If Psw = "" Then
         Psw = GetConnectInfo(lDbRemu, "PWD")
         UsrPsw = "User=" & Usr & ";PWD=" & Psw & ";"
      End If
      
'      gConnStr = ConnStr & UsrPsw   ' Para la exportación
      
'      lDbRemuDate = GetDbNow(lDbRemu)
      
   End If

End Function

