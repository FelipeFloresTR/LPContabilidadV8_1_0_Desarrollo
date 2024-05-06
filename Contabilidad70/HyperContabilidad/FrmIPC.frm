VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmIPC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantención de Valores e Índices"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "FrmIPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_FactoresAnuales 
      Caption         =   "Factores Anuales..."
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      ToolTipText     =   "Obtener Punto IPC del mes seleccionado"
      Top             =   5280
      Width           =   1515
   End
   Begin VB.CommandButton Bt_IPC 
      Caption         =   "Obtener IPC"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "Obtener Punto IPC del mes seleccionado"
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1500
      TabIndex        =   4
      Top             =   480
      Width           =   3915
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   480
      Picture         =   "FrmIPC.frx":000C
      ScaleHeight     =   555
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   660
      Width           =   615
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5700
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5700
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   3435
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6059
      Cols            =   8
      Rows            =   15
      FixedCols       =   0
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
End
Attribute VB_Name = "FrmIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_HMES = 0
Const C_MES = 1
Const C_PIPC = 2
Const C_VIPC = 3
Const C_AIPC = 4
Const C_FCM = 5
Const C_ID = 6
Const C_ST = 7

Dim PIpc(20) As Double


Private Sub SetupForm()
   Dim c As Integer

   Grid.ColWidth(C_HMES) = 0
   
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ST) = 0
   Grid.ColWidth(C_PIPC) = 1400
   
   Grid.TextMatrix(0, C_MES) = "Mes"
   Grid.TextMatrix(0, C_PIPC) = "Ptos. IPC"
   Grid.TextMatrix(0, C_VIPC) = "Var. IPC %"
   Grid.TextMatrix(0, C_AIPC) = "IPC Acum."
   Grid.TextMatrix(0, C_FCM) = "Fact. Act."
   
   For c = C_PIPC To C_FCM
      Grid.ColAlignment(c) = flexAlignRightCenter
   Next c


End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim m0 As Integer, a As Integer, m As Integer, Dt1 As Long, Dt2 As Long
   Dim x As Integer, r As Integer, lr As Integer, i As Integer
   Dim aIPC As Double
   Dim Año As Integer
   Dim c As Integer
   
   For i = 0 To UBound(PIpc)
      PIpc(i) = 0
   Next i
   
   Año = CbItemData(Cb_Ano)

   Dt1 = DateSerial(Año - 1, 11, 1)
   Dt2 = DateSerial(Año, 12, 1)
   
   For r = Grid.FixedRows To Grid.rows - 1
      For c = 0 To Grid.Cols - 1
         Grid.TextMatrix(r, c) = ""
      Next c
   Next r
   
   Q1 = "SELECT AnoMes, PIPC, VIPC, aIPC, FCM FROM IPC WHERE AnoMes BETWEEN " & Dt1 & " AND " & Dt2 & " ORDER BY AnoMes"
   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows + 14
   
   lr = Grid.FixedRows - 1
   aIPC = 0
   r = Grid.FixedRows
   
   Do Until Rs.EOF
       
      a = Year(vFld(Rs("AnoMes")))
      m = Month(vFld(Rs("AnoMes")))
      r = ((a - Año) * 12 + m) + 2
         
      If r > lr + 1 Then
         For i = lr + 1 To r - 1
         
            Dt1 = DateSerial(Año - 1, 10 + i, 1)
            Grid.TextMatrix(i, C_HMES) = Dt1
            Grid.TextMatrix(i, C_MES) = Left(gNomMes(Month(Dt1)), 3) & " " & Year(Dt1)
         Next i

      End If
      
      lr = r
      
      PIpc(r) = vFld(Rs("pIPC"))
            
      Grid.TextMatrix(r, C_ID) = vFld(Rs("AnoMes"))
      Grid.TextMatrix(r, C_HMES) = vFld(Rs("AnoMes"))
      
      Grid.TextMatrix(r, C_MES) = Left(gNomMes(m), 3) & " " & a
      
      If PIpc(r) Then
         If Año > 2000 Then
            
            If Año = 2019 Then         'Diconsinuidad del INE (Victor Morales, 20 ago 2019)
               If a = 2018 And m = 12 Then        'Dic 2018
                  PIpc(r) = 100.64
               End If
            End If
            
             Grid.TextMatrix(r, C_PIPC) = Format(PIpc(r), DBLFMT3)
         Else
            Grid.TextMatrix(r, C_PIPC) = Format(PIpc(r), DBLFMT4)
         End If
         

         Grid.TextMatrix(r, C_VIPC) = Format(vFld(Rs("vIPC")) * 100, DBLFMT1)
         
         If a <> Año - 1 Then  'es dic año anterior
            Grid.TextMatrix(r, C_AIPC) = Format(vFld(Rs("aIPC")) * 100, DBLFMT1)
         End If
      End If
      
'      aIPC = aIPC + pIPC(r)      'en Recalc se calcula de otra manera
'      If r > 1 Then
'         Grid.TextMatrix(r, C_AIPC) = Format(aIPC * 100, DBLFMT1)
'      End If
'

      If Trim(Grid.TextMatrix(r, C_PIPC)) <> "" Then
         Grid.TextMatrix(r, C_FCM) = Format(vFld(Rs("fCM")), DBLFMT3)
      End If

      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   For i = r + 1 To Grid.rows - 1
      Dt1 = DateSerial(Año - 1, 10 + i, 1)
      Grid.TextMatrix(i, C_HMES) = Dt1
      Grid.TextMatrix(i, C_MES) = Left(gNomMes(Month(Dt1)), 3) & " " & Year(Dt1)
   Next i

   Grid.RowHeight(1) = 0

   If Año <= 2018 Or (Año = 2019 And vFmt(Grid.TextMatrix(Grid.FixedRows + 5, C_AIPC)) = 0) Then
      Call Recalc
   End If
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_FactoresAnuales_Click()
   Dim Frm As FrmFactoresAct
   
   Set Frm = New FrmFactoresAct
   Call Frm.FEdit(Val(Cb_Ano))
   Set Frm = Nothing
   
End Sub

Private Sub Bt_IPC_Click()
   Dim Value As Double, r As Integer, AnoMes As Long, Ano As Integer, Mes As Integer
   Dim AnoSeleccionado As Integer
   Dim SII_IPCs(12) As SII_IPC_t, SII_Fact(12) As SII_Fact_t
   Dim Rc As Integer, m As Integer
   
   r = Grid.Row
   If r = 1 Then
      r = r + 1
   End If
   AnoMes = Val(Grid.TextMatrix(r, C_HMES))
   If AnoMes <= 0 Then
      Exit Sub
   End If
   
   Ano = Year(AnoMes)
   Mes = Month(AnoMes)
   
   AnoSeleccionado = CbItemData(Cb_Ano)

   If AnoSeleccionado <= 2012 Then
      Value = LPGetValorMes("IPC", Ano, Mes)
      If Value <> -7777 Then
      
         If AnoSeleccionado = 2019 Then         'Diconsinuidad del INE (Victor Morales, 20 ago 2019)
            If Ano = 2018 And Mes = 12 Then        'Dic 2018
               Value = 100.64
            End If
         End If
   
         If PIpc(r) <> Value Then
            PIpc(r) = Value
            Grid.TextMatrix(r, C_PIPC) = Format(PIpc(r), DBLFMT3)
            Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
            Cb_Ano.Locked = True
            Call Recalc
         End If
   
      End If
   
   Else
   
      Rc = SII_GetIPCs(AnoSeleccionado, SII_IPCs(), False)
      If Rc <> 0 Then
         MsgBox1 "Error al obtener los valores de IPC desde el sitio del SII.", vbExclamation
         Exit Sub
      End If
      
      r = Grid.FixedRows + 2
      For m = 1 To 12
      
         Grid.TextMatrix(r, C_PIPC) = IIf(SII_IPCs(m).bPIpc, Format(SII_IPCs(m).PIpc, DBLFMT3), "")
         Grid.TextMatrix(r, C_VIPC) = IIf(SII_IPCs(m).bVarIpc, Format(SII_IPCs(m).VarIpc, DBLFMT1), "")
         Grid.TextMatrix(r, C_AIPC) = IIf(SII_IPCs(m).bVarAcum, Format(SII_IPCs(m).VarAcum, DBLFMT1), "")
         Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)

         r = r + 1
      Next m
               
      Rc = SII_CorrMonet(AnoSeleccionado, SII_Fact())
      If Rc <> 0 Then
         MsgBox1 "Error al obtener los valores de los factores desde el sitio del SII.", vbExclamation
         Exit Sub
      End If
      
      r = Grid.FixedRows + 2
      For m = 1 To 12
      
         Grid.TextMatrix(r, C_FCM) = IIf(SII_Fact(m).bFact, Format(SII_Fact(m).Fact, DBLFMT3), "")
         
         r = r + 1
      Next m
               
   End If
   
   Cb_Ano.Enabled = False

End Sub

Private Sub Bt_OK_Click()

   Call SaveAll
   
   Unload Me

End Sub

Private Sub Cb_Ano_Click()

'   If Val(Cb_Ano) >= gEmpresa.Ano Then
'      Bt_OK.Enabled = True
'   Else
'      Bt_OK.Enabled = False
'   End If
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   Dim Ano As Integer
   
   Call SetupForm
   
   Ano = gEmpresa.Ano
   If Ano <= 0 Then
      Ano = Year(Now)
   End If
   
   Call CbFillAno(Cb_Ano, Ano, 2000, Year(Now) + 1)
   
   Call SetupPriv

   'Call LoadAll
   
End Sub

Private Sub Recalc()
   Dim r As Integer, aIPC As Double, vIPC As Double, fCM As Double, UltIPC As Double, b As Boolean
   Dim rUlt As Integer, r1 As Integer, Dt As Long, Ano As Integer
   Dim Valor As String

   UltIPC = 0
   b = False

   Ano = CbItemData(Cb_Ano)

   ' buscamos el más reciente - 1
   For r = Grid.rows - 1 To Grid.FixedRows Step -1
      If PIpc(r) Then
         UltIPC = PIpc(r)
         rUlt = r
         If r = Grid.rows - 2 Then    'es noviembre, entregamos los factores de diciembre, dado que se usa el IPC de noviembre
            Exit For
         End If
         If b Then
            Exit For
         End If
         b = True
      End If
   Next r
   
   r1 = Grid.FixedRows
   
   For r = r1 To rUlt + 1
   
      ' calculamos el % de IPC
      If r <= r1 Then
         aIPC = 0
      Else
         If PIpc(r - 1) <> 0 And PIpc(r) <> 0 Then
            vIPC = (PIpc(r) - PIpc(r - 1)) / PIpc(r - 1)
            Dt = Val(Grid.TextMatrix(r, C_HMES))
            
            If r < Grid.rows Then
               If Ano = 2019 And Dt = CLng(DateSerial(2018, 12, 1)) Then   'dic 2018
                  vIPC = 0.005       'discontinuidad generada por el INE, por cambio de base (Victor Morales 5 mar 2020)
               End If
               If Dt = CLng(DateSerial(2019, 1, 1)) Then   'enero 2019
                  vIPC = 0.001       'discontinuidad generada por el INE, por cambio de base (Joshua Nicolás Catrin)
               End If
            End If

            If Year(Dt) = Ano Then  ' desde Enero
               aIPC = aIPC + vIPC
            End If

'            Valor = Format(vIPC * 100, DBLFMT1)
            Valor = Round(vIPC * 100, 3)     'más decimales para ser más certero con lo del SII
            Valor = Round(Abs(Valor), 1) * Sgn(vIPC)
                        
            If Grid.TextMatrix(r, C_VIPC) <> Valor Then
               Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
               Grid.TextMatrix(r, C_VIPC) = Format(Valor, DBLFMT1)
            End If
           
            ' el acumulado
            If Year(Dt) = Ano Then  ' desde Enero
'               Valor = Format(aIPC * 100, DBLFMT1)
'               Valor = Round(aIPC * 100, 1)
               Valor = Round(aIPC * 100, 3)     'más decimales para ser más certero con lo del SII
               Valor = Round(Abs(Valor), 1) * Sgn(aIPC)
               
               If Ano = 2019 And Dt = CLng(DateSerial(2018, 12, 1)) Then   'dic 2018
                  Valor = 2.6      'discontinuidad generada por el INE, por cambio de base (Victor Morales 5 mar 2020)
               End If

               If Dt = CLng(DateSerial(2019, 2, 1)) Then   'feb 2019
                  Valor = 0.2       'discontinuidad generada por el INE, por cambio de base (Victor Morales)
               End If

               If Dt = CLng(DateSerial(2019, 6, 1)) Then   'junio 2019
                  Valor = 1.6       'discontinuidad generada por el INE, por cambio de base (Victor Morales)
               End If
               
               If Dt = CLng(DateSerial(2019, 8, 1)) Then   'ago 2019
                  Valor = 2           'discontinuidad generada por el INE, por cambio de base (Victor Morales)
               End If
               
               If Dt = CLng(DateSerial(2019, 11, 1)) Then   'nov 2019
                  Valor = 2.9        'discontinuidad generada por el INE, por cambio de base (Victor Morales)
               End If
               
               If Grid.TextMatrix(r, C_AIPC) <> Valor Then
                  Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
                  Grid.TextMatrix(r, C_AIPC) = Format(Valor, DBLFMT1)
               End If
            
            ElseIf Ano = 2019 And Dt = CLng(DateSerial(2018, 12, 1)) Then   'dic 2018
               Valor = 2.6      'discontinuidad generada por el INE, por cambio de base (Victor Morales 5 mar 2020)
               Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
               Grid.TextMatrix(r, C_AIPC) = Format(Valor, DBLFMT1)
            End If
            
            ' el acumulado
'            If pIPC(r1 + 1) <> 0 Then
'               aIPC = (pIPC(r) - pIPC(r1 + 1)) / pIPC(r1 + 1)
''              aIPC = aIPC + vIPC
'               Valor = Format(aIPC * 100, DBLFMT1)
'            Else
'               'MsgBox1 "En la fila " & r1 + 1 & " falta ingresar valor de IPC. El cálculo quedará incompleto.", vbExclamation
'               Valor = ""
'            End If
'
'            If r <= rUlt + 1 Then  ' por algún motivo va corrido en uno
'               If Grid.TextMatrix(r, C_AIPC) <> Valor Then
'                  'Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
'                  Grid.TextMatrix(r, C_AIPC) = Valor
'               End If
'            End If
            
                      
         Else
            If r < Grid.rows Then
               If Grid.TextMatrix(r, C_VIPC) <> "" Or Grid.TextMatrix(r, C_AIPC) <> "" Or Grid.TextMatrix(r, C_FCM) <> "" Then
                  Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
                  Grid.TextMatrix(r, C_VIPC) = ""
                  Grid.TextMatrix(r, C_AIPC) = ""
                  Grid.TextMatrix(r, C_FCM) = ""
               End If
            End If
         End If
      End If
      
      ' calculamos el factor de actualizacion
      If PIpc(r - 1) Then
         
         fCM = UltIPC / PIpc(r - 1)

         If fCM < 1 Then
            fCM = 1
         End If
         
         If r = Grid.FixedRows + 1 Then
            If Grid.TextMatrix(r, C_HMES) = CLng(DateSerial(2009, 12, 1)) Then    'dic 2009
               fCM = 1.025       'discontinuidad generada por el INE
            End If
         End If
         
         If r < Grid.rows Then
            If Grid.TextMatrix(r, C_HMES) = CLng(DateSerial(2010, 1, 1)) Then   'enero 2010
               fCM = 1.029       'discontinuidad generada por el INE
            End If
         End If
         
         If r = Grid.FixedRows + 1 Then   'primer registro
            If Grid.TextMatrix(r, C_HMES) = CLng(DateSerial(2018, 12, 1)) Then    'dic 2018, habiendo seleccionado año 2019
               fCM = 1.028       'discontinuidad generada por el INE (Victor Morales 5 mar 2020
            End If
         End If
     
         Valor = Format(fCM, DBLFMT3)
         If r < Grid.rows Then
            If Grid.TextMatrix(r, C_FCM) <> Valor Then
               Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
               Grid.TextMatrix(r, C_FCM) = Valor
            End If
         End If
      Else
         
      End If

   Next r

End Sub
Private Sub RecalcAcum()
   Dim i As Integer
   
   For i = Grid.FixedRows + 3 To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_VIPC) <> "" Then
         Grid.TextMatrix(i, C_AIPC) = Format(vFmt(Grid.TextMatrix(i - 1, C_AIPC)) + vFmt(Grid.TextMatrix(i, C_VIPC)), DBLFMT1)
      Else
         Grid.TextMatrix(i, C_AIPC) = ""
      End If
         
   Next i

End Sub
Private Sub Recalc_old()
   Dim r As Integer, aIPC As Double, vIPC As Double, fCM As Double, UltIPC As Double, b As Boolean
   Dim rUlt As Integer, r1 As Integer
   Dim Valor As String

   UltIPC = 0
   b = False

   ' buscamos el más reciente - 1
   For r = Grid.rows - 1 To Grid.FixedRows Step -1
      If PIpc(r) Then
         UltIPC = PIpc(r)
         rUlt = r
         If r = Grid.rows - 2 Then    'es noviembre, entregamos los factores de diciembre, dado que se usa el IPC de noviembre
            Exit For
         End If
         If b Then
            Exit For
         End If
         b = True
      End If
   Next r
   
   r1 = Grid.FixedRows
   
   For r = r1 To rUlt + 1
   
      ' calculamos el % de IPC
      If r <= r1 Then
         aIPC = 0
      Else
         If PIpc(r - 1) <> 0 And PIpc(r) <> 0 Then
            vIPC = (PIpc(r) - PIpc(r - 1)) / PIpc(r - 1)
            
            If r < Grid.rows Then
               If Val(Grid.TextMatrix(r, C_HMES)) = CLng(DateSerial(2019, 1, 1)) Then   'enero 2019
                  vIPC = 0.001       'discontinuidad generada por el INE, por cambio de base
               End If
            End If

            Valor = Format(vIPC * 100, DBLFMT1)
            Valor = Round(vIPC * 100, 2)
            Valor = Round(Abs(Valor), 1) * Sgn(vIPC)
            
            If Grid.TextMatrix(r, C_VIPC) <> Valor Then
               Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
               Grid.TextMatrix(r, C_VIPC) = Valor
            End If
           
            ' el acumulado
            If PIpc(r1 + 1) <> 0 Then
               aIPC = (PIpc(r) - PIpc(r1 + 1)) / PIpc(r1 + 1)
'              aIPC = aIPC + vIPC
               Valor = Format(aIPC * 100, DBLFMT1)
            Else
               'MsgBox1 "En la fila " & r1 + 1 & " falta ingresar valor de IPC. El cálculo quedará incompleto.", vbExclamation
               Valor = ""
            End If
            If r <= rUlt + 1 Then  ' por algún motivo va corrido en uno
               If Grid.TextMatrix(r, C_AIPC) <> Valor Then
                  'Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
                  Grid.TextMatrix(r, C_AIPC) = Valor
               End If
            End If
            
                      
         Else
            If r < Grid.rows Then
               If Grid.TextMatrix(r, C_VIPC) <> "" Or Grid.TextMatrix(r, C_AIPC) <> "" Or Grid.TextMatrix(r, C_FCM) <> "" Then
                  Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
                  Grid.TextMatrix(r, C_VIPC) = ""
                  Grid.TextMatrix(r, C_AIPC) = ""
                  Grid.TextMatrix(r, C_FCM) = ""
               End If
            End If
         End If
      End If
      
      ' calculamos el factor de actualizacion
      If PIpc(r - 1) Then
         
         fCM = UltIPC / PIpc(r - 1)

         If fCM < 1 Then
            fCM = 1
         End If
         
         If r = Grid.FixedRows + 1 Then
            If Grid.TextMatrix(r, C_HMES) = CLng(DateSerial(2009, 12, 1)) Then    'dic 2009
               fCM = 1.025       'discontinuidad generada por el INE
            End If
         End If
         
         If r < Grid.rows Then
            If Grid.TextMatrix(r, C_HMES) = CLng(DateSerial(2010, 1, 1)) Then   'enero 2010
               fCM = 1.029       'discontinuidad generada por el INE
            End If
         End If
      
         Valor = Format(fCM, DBLFMT3)
         If r < Grid.rows Then
            If Grid.TextMatrix(r, C_FCM) <> Valor Then
               Call FGrModRow(Grid, r, FGR_U, C_ID, C_ST)
               Grid.TextMatrix(r, C_FCM) = Valor
            End If
         End If
      End If

   Next r

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Col <> C_PIPC Then
      Exit Sub
   End If

   PIpc(Row) = vFmt(Value)
   
   Value = Format(PIpc(Row), DBLFMT3)
   If Value <> Grid.TextMatrix(Row, C_PIPC) Then
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_ST)
      Cb_Ano.Locked = True
      Call Recalc
   End If
   
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)

   'If Col <> C_PIPC Or Val(Cb_Ano) < gEmpresa.Ano Then
   If Col <> C_PIPC Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   EdType = FEG_Edit

End Sub

Private Sub SaveAll()
   Dim r As Integer, Q1 As String, Rc As Long

   Me.MousePointer = vbHourglass
   For r = Grid.FixedRows To Grid.rows - 1

      If Grid.TextMatrix(r, C_ST) = FGR_U Then
         Q1 = "UPDATE IPC SET pIPC=" & str(vFmt(Grid.TextMatrix(r, C_PIPC)))
         Q1 = Q1 & ", vIPC=" & str(vFmt(Grid.TextMatrix(r, C_VIPC)) / 100)
         Q1 = Q1 & ", aIPC=" & str(vFmt(Grid.TextMatrix(r, C_AIPC)) / 100)
         Q1 = Q1 & ", fCM=" & str(vFmt(Grid.TextMatrix(r, C_FCM)))
         Q1 = Q1 & " WHERE AnoMes=" & Grid.TextMatrix(r, C_HMES)
         Rc = ExecSQL(DbMain, Q1)

      ElseIf Grid.TextMatrix(r, C_ST) = FGR_I Then
         Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, aIPC, fCM)"
         Q1 = Q1 & " VALUES (" & Grid.TextMatrix(r, C_HMES)
         Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(r, C_PIPC)))
         Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(r, C_VIPC)) / 100)
         Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(r, C_AIPC)) / 100)
         Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(r, C_FCM)))
         Q1 = Q1 & " )"
         Rc = ExecSQL(DbMain, Q1)
      End If
      
   Next r

   're-leemos los datos a memoria
   Call ReadIndices
   Me.MousePointer = vbDefault

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, "Indices")
   End If
End Sub
Private Sub SetupPriv()

   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
      Cb_Ano.Locked = False
   End If

End Sub

