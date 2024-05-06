VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCierreAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre Período"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   Icon            =   "FrmCierreAnual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1320
      TabIndex        =   2
      Top             =   300
      Width           =   8475
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   7935
         Begin MSComctlLib.ProgressBar ProgressBar 
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Min             =   1e-4
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre año"
         Height          =   255
         Index           =   0
         Left            =   3660
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "2005"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   10140
      TabIndex        =   1
      Top             =   840
      Width           =   1395
   End
   Begin VB.CommandButton Bt_CerrarAno 
      Caption         =   "Cerrar Período"
      Height          =   315
      Left            =   10140
      TabIndex        =   0
      Top             =   420
      Width           =   1395
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   360
      Picture         =   "FrmCierreAnual.frx":000C
      Top             =   420
      Width           =   780
   End
End
Attribute VB_Name = "FrmCierreAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'3389677 FP
Dim remMesAnt As Boolean

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CerrarAno_Click()
   Dim Msg As String
   Dim Q1 As String
   Dim i As Integer
   Dim Rs As Recordset
   Dim MaxCorr As Long
   Dim LastComp(N_TIPOCOMP) As Long
   Dim F1 As Long
   Dim Rc As Integer
   Dim RemIVAUTM As Double
   Dim SaldoLibroCaja As Double
   
   'If gEmpresa.FApertura = 0 Then
   '   MsgBox1 "No puede hacer cierre año " & gEmpresa.Ano & " ya que no ha hecho la apertura año " & gEmpresa.Ano + 1 & ".", vbExclamation
   '   Exit Sub
   'End If
   
   If LibAnualesImpresos(True) = False Then
      Exit Sub
   End If
   
   Msg = "Recuerde que al hacer cierre anual no podrá volver a modificar datos de este año. ¿Desea continuar?"
   If MsgBox1(Msg, vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   'obtenemos último correlativo comprobante
   If gPerCorrComp = TCC_CONTINUO Then
      If gTipoCorrComp = TCC_UNICO Then
         Q1 = "SELECT Max(Correlativo) FROM Comprobante WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            MaxCorr = vFld(Rs(0))
         End If
         
         Call CloseRs(Rs)
      Else
         
         Q1 = "SELECT Tipo, Max(Correlativo) as LastComp FROM Comprobante WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " GROUP BY Tipo "
         Set Rs = OpenRs(DbMain, Q1)
         
         Do While Rs.EOF = False
            LastComp(vFld(Rs("Tipo"))) = vFld(Rs("LastComp"))
            Rs.MoveNext
         Loop
         
         Call CloseRs(Rs)
         
      End If
   End If
   
   'cerramos todos los meses abiertos
   Set Rs = OpenRs(DbMain, "SELECT Mes, Estado FROM EstadoMes WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
   Do While Rs.EOF = False
      
      If vFld(Rs("Estado")) = EM_ABIERTO Then
         Call CerrarMes(vFld(Rs("Mes")))
      End If
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   '3340329 empieza solucion
   Dim vTotalRemMesAnt As Double
   Dim TotIVACred As Double
   Dim TotIVADeb As Double
   Dim TotRemUTM As Double
   Dim IVAIrrec As Double, IVARetParcial As Double, IVARetTotal As Double
   Dim TotIEPDGen As Double, TotIEPDTransp As Double
   Dim Fecha As Double
   Dim ValUTM As Double
   Dim RemMesAntUTM As Double
   Dim TotRemMesAnt As Double
   Dim Mes As Double
       
   For i = 1 To 13 - 1
   
   '3389677 FPR se creo para que no traspase si es un saldo a pagar al mes siguiente, ya que no es remanente
   
   Mes = month(DateSerial(Val(gEmpresa.Ano), i, 1))
   '3426794 ffv
'   If i = Mes Then
'        If Not remMesAnt Then
'            'Tx_RemMesAnt.Text = 0
'            TotRemMesAnt = 0
'            vTotalRemMesAnt = 0
'            'TotIVA = 0
'        End If
'   End If
   
    If Not Mes = 1 And RemIVAAnoAnt = True Then
        If Not remMesAnt Then
            'Tx_RemMesAnt.Text = 0
            TotRemMesAnt = 0
            vTotalRemMesAnt = 0
        End If
     End If
    '3426794 ffv
   'FIN 3389677 FPR
   
    Rc = GetRemIVAUTM_New(i, gEmpresa.Ano, RemIVAUTM, vFmt(Format(Abs(vTotalRemMesAnt), NEGNUMFMT)))
    If Rc < 0 Then
       RemIVAUTM = 0
       If Rc = ERR_VALUTM Then
          MsgBox1 "No se encontró el valor de la UTM para calcular Remanente Crédito IVA para año siguiente.", vbExclamation
       End If
    End If
    
     Fecha = DateSerial(Val(gEmpresa.Ano), i + 1, 1)     'se agrega + 1 a solicitud de Victor Morales (17 nov. 2011)
    
   If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
      TotRemMesAnt = RemIVAUTM * ValUTM
   Else
      TotRemMesAnt = 0
   End If

   Call GetResIVA(i, Val(gEmpresa.Ano), TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
   '   '3340329
   
   TotIVACred = TotIVACred - IVAIrrec          'se resta IVA Irrecuperable que se obtiene en la función LoadValOImp
      
   TotIVADeb = TotIVADeb - IVARetParcial - IVARetTotal        'se resta IVA Retenido Parcial o Total que se obtiene en la función LoadValOImp
    
   Dim AjusteIvaMen As Double
   'Dim Mes As Integer
   'Mes = ItemData(Cb_Mes)
   AjusteIvaMen = GetAjusteIVAMensual(i)
    
   vTotalRemMesAnt = TotIVADeb - (TotIVACred + TotIEPDGen + TotIEPDTransp + TotRemMesAnt + AjusteIvaMen)
   
   
   If vTotalRemMesAnt < 0 Then
      remMesAnt = True
   Else
   remMesAnt = False
   End If
   '3389677 FPR
   
   Next i
   
   If ValUTM = 0 Then
       RemIVAUTM = 0
       MsgBox1 "No se encontró el valor de la UTM para calcular Remanente Crédito IVA para año siguiente.", vbExclamation
   Else
        RemIVAUTM = Format(vFmt(Format(Abs(vTotalRemMesAnt), NEGNUMFMT)) / ValUTM, DBLFMT2)
   End If
   
   'RemIVAUTM = Format(vFmt(Format(Abs(vTotalRemMesAnt), NEGNUMFMT)) / ValUTM, DBLFMT2)
   '3340329 termina solucion requerimiento
   
   '3340329 se comenta para calcular remantente en base a total mes anterior
   'almacenamos el remanente de IVA de este año para utilizarlo el año siguiente (se llama con un mes más que diciembre para que incluya diciembre)
'   Rc = GetRemIVAUTM(12 + 1, gEmpresa.Ano, RemIVAUTM)
'   If Rc < 0 Then
'      RemIVAUTM = 0
'      If Rc = ERR_VALUTM Then
'         MsgBox1 "No se encontró el valor de la UTM para calcular Remanente Crédito IVA para año siguiente.", vbExclamation
'      End If
'   End If
    '3340329
   
   'almacenamos el saldo final del libro de caja, para utilizarlo el año siguiente
   SaldoLibroCaja = 0
   Q1 = "SELECT Sum(Ingreso) as Ingresos, Sum(Egreso) as Egresos FROM LibroCaja "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      SaldoLibroCaja = vFld(Rs("Ingresos")) - vFld(Rs("Egresos"))
   End If
   Call CloseRs(Rs)
   
   
   'actualizamos fecha de cierre, correlativos y remanente de IVA Crédito
   F1 = CLng(Int(Now))
   
   
   Q1 = "UPDATE EmpresasAno SET FCierre=" & F1
   Q1 = Q1 & ", NumLastCompUnico=" & MaxCorr
   
   For i = 1 To N_TIPOCOMP
      Q1 = Q1 & ", NumLastComp" & Left(gTipoComp(i), 1) & "=" & LastComp(i)
   Next i
   
   Q1 = Q1 & ", RemIVAUTM = " & str(RemIVAUTM)
   Q1 = Q1 & ", SaldoLibroCaja = " & SaldoLibroCaja
   
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id & " AND Ano=" & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   gEmpresa.FCierre = F1
     
   For i = 1 To 100
      Sleep (10)
      ProgressBar.Value = i
   Next i
      
   Me.MousePointer = vbDefault
   
   Unload Me
End Sub

Private Sub Form_Load()
   Label1(1) = gEmpresa.Ano
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv
   
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_EMPRESA) Then
      Call EnableForm(Me, False)
   End If
   
End Function
