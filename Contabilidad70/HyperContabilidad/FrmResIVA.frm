VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen IVA Compras - Ventas"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "FrmResIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Otros Impuestos"
      Top             =   5280
      Width           =   8055
   End
   Begin VB.Frame Frame2 
      Caption         =   "IVA"
      Height          =   3795
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   4815
      Begin VB.TextBox Tx_AjusRemCRedFis 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2160
         Width           =   1515
      End
      Begin VB.TextBox Tx_IEPDTrCF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1740
         Width           =   1515
      End
      Begin VB.TextBox Tx_IEPDGenCF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox Tx_TotalUTM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox Tx_RemMesAnt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox Tx_IVACred 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1020
         Width           =   1515
      End
      Begin VB.TextBox Tx_IVADeb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox Tx_Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   13
         Left            =   2760
         TabIndex        =   34
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ajuste Remanente Crédito Fiscal:"
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   32
         Top             =   2160
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total IEPD Transporte CF:"
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   31
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   10
         Left            =   2820
         TabIndex        =   30
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total IEPD General CF:"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   8
         Left            =   2820
         TabIndex        =   27
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Lb_RemUTM 
         AutoSize        =   -1  'True
         Caption         =   "UTM"
         Height          =   195
         Left            =   2520
         TabIndex        =   24
         Top             =   3240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Lb_TotalUTM 
         Caption         =   "Remanente IVA Crédito Fiscal"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   3240
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   21
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   6
         Left            =   2820
         TabIndex        =   20
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   5
         Left            =   2820
         TabIndex        =   19
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Index           =   2
         Left            =   2820
         TabIndex        =   18
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label1 
         Caption         =   "Remanente periodo anterior:"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Total Crédito Fiscal:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Total Débito Fiscal:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Lb_Total 
         Caption         =   "Remanente periodo siguiente:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   4500
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.CommandButton Bt_ResLibAux 
      Caption         =   "Resumen Libros      Auxiliares..."
      Height          =   1035
      Left            =   7140
      Picture         =   "FrmResIVA.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7140
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   4815
      Begin VB.TextBox Tx_Ano 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1515
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   420
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2955
      Left            =   360
      TabIndex        =   35
      Top             =   5640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Label Lb_DefCta 
      Caption         =   "Falta definir cuenta de Crédito IVA en config. cuentas básicas"
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   5340
      TabIndex        =   25
      Top             =   2220
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label Lb_UTM 
      Caption         =   "Falta valor UTM para calcular remanente"
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   5340
      TabIndex        =   16
      Top             =   1620
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "FrmResIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_TIPOLIB = 0
Const C_CODVALLIB = 1
Const C_LIBRO = 2
Const C_DESC = 3
Const C_RECUPERABLE = 4
Const C_VALOR = 5

Const NCOLS = C_VALOR

Dim lMes As Integer
Dim lAno As Integer
Dim lVerBotonRes As Boolean
Dim lTipoLib As Integer

Dim lIVAIrrec As Double
Dim lIVARetParcial As Double
Dim lIVARetTotal As Double
Dim AjusRemCRedFis As Double 'gcb21092021
Dim Fecha As Double
Dim ValUTM As Double
Dim lInLoad As Boolean

Dim lModVal As Boolean        'FCA 23 sep 2021

'3389677 FPR
Dim remMesAnt As Boolean


Private Sub bt_Cerrar_Click()

   If vFmt(Tx_IEPDGenCF) <> 0 Or vFmt(Tx_IEPDTrCF) <> 0 Then
      MsgBox1 "Recuerde que para realizar la recuperación del Impuesto Específico al Petróleo Diesel General o Trasportistas de carga debe cumplir con los requisitos establecidos en el art. 7°, de la Ley 18.502, Arts. 1° y 3° D.S. N° 311, de 1986 y el art. 2° de la Ley N° 19.764.", vbInformation
   End If
   Unload Me
End Sub

Private Sub Bt_OK_Click()        'FCA 23 sep 2021
   
   Call SaveAll
   
   Unload Me
   
End Sub

Private Sub Bt_ResLibAux_Click()
   Dim Frm As FrmResLibAux
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResLibAux
   Call Frm.FView(lMes)
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub Cb_Mes_Click()

   Me.MousePointer = vbHourglass
   
   If lModVal Then         'FCA   23 sep 2021
      If MsgBox1("¿Desea almacenar el valor del Ajuste Remanente Crédito Fiscal?", vbYesNo) = vbYes Then
         Call SaveAll
      End If
   End If
   
   Call LoadValOImp   'este debe ir antes de LoadVal porque obtienen valor de IVA Irrecuperable que se usa en LoadVal
   Call LoadVal
   
   Me.MousePointer = vbDefault
   
   lMes = CbItemData(Cb_Mes)

End Sub

Private Sub SaveAll()         'FCA 23 sep 2021

   On Error GoTo ERR
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rc As Long
   Dim Mes As Integer

'   Mes = ItemData(Cb_Mes)
   Q1 = "SELECT * FROM AjusteIVAMensual WHERE Ano=" & lAno & " And idempresa= " & gEmpresa.id & " And Mes=" & lMes & ""
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      Q1 = "UPDATE AjusteIVAMensual SET VALOR=" & vFmt(Tx_AjusRemCRedFis) & ""
      Q1 = Q1 & " WHERE Ano=" & lAno & " And idempresa= " & gEmpresa.id & " And Mes=" & lMes & ""
      Rc = ExecSQL(DbMain, Q1)
   Else
      
      Q1 = "INSERT INTO AjusteIVAMensual (idEmpresa, Ano, Mes, Valor) VALUES (" & gEmpresa.id & " ," & lAno & "," & lMes & "," & vFmt(Tx_AjusRemCRedFis) & ")"
      Rc = ExecSQL(DbMain, Q1)
   End If
   Call CloseRs(Rs)
   
   lModVal = False

Exit Sub
ERR:
   If ERR.Number <> 0 Then
     MsgBox ERR.Description
   End If
End Sub
Sub Calcular()      ' FCA 23 sep 2021
   Dim Total As Double
   Dim ValUTM As Double
   Dim TotRemUTM As Double
   Dim Fecha As Long
   
   Fecha = DateSerial(gEmpresa.Ano, ItemData(Cb_Mes) + 1, 1)
   
   If GetValMoneda("UTM", ValUTM, Fecha, True) = False Then   'no se puede calcular
      Tx_TotalUTM = 0
      Exit Sub
   End If

   Total = vFmt(Tx_IVADeb) - (vFmt(Tx_IVACred) + vFmt(Tx_IEPDGenCF) + vFmt(Tx_IEPDTrCF) + vFmt(Tx_RemMesAnt) + vFmt(Tx_AjusRemCRedFis))
   Tx_Total = Format(Abs(Total), NEGNUMFMT)
   
   If Total < 0 Then   'remanente
      Lb_Total = "Remanente periodo siguiente:"
      Lb_TotalUTM.visible = True
      Tx_TotalUTM.visible = True
      Lb_RemUTM.visible = True
      
      If Lb_UTM.visible = False Then
         If ValUTM <> 0 Then
            TotRemUTM = vFmt(Tx_Total) / ValUTM
         End If
         Tx_TotalUTM = Format(TotRemUTM, DBLFMT2)
      End If
      
      remMesAnt = True
      
   Else                          'a pagar
      Lb_Total = "Sub Total a Pagar:"
      Lb_TotalUTM.visible = False
      Tx_TotalUTM.visible = False
      Lb_RemUTM.visible = False
      
      remMesAnt = False
      
   End If

End Sub
Private Sub Form_Load()
   Dim MesActual As Integer
   
   lInLoad = True
   
   Call SetUpGrid
   Tx_Ano = lAno

   MesActual = GetMesActual()
   
   Cb_Mes.AddItem " "
   Cb_Mes.ItemData(Cb_Mes.NewIndex) = 0
    
   Call FillMes(Cb_Mes)
               
   If lMes > 0 Then
      Cb_Mes.ListIndex = lMes
   Else
      If MesActual > 0 Then
         Cb_Mes.ListIndex = MesActual
      Else
         Cb_Mes.ListIndex = GetUltimoMesConMovs()
      End If
   End If
   
   If Not lVerBotonRes Then
      Bt_ResLibAux.visible = False
   End If
   
   lInLoad = False
      
End Sub

Private Sub LoadVal()
   Dim TotIVACred As Double
   Dim TotIVADeb As Double
   Dim TotRemMesAnt As Double
   Dim Mes As Integer
   Dim RemMesAntUTM As Double
   
   
   Dim TotRemUTM As Double
   Dim Rc As Integer
   Dim TotIEPDGen As Double, TotIEPDTransp As Double
   Dim AjusteIVAMensual As Double
   Dim i As Integer
   
   Mes = ItemData(Cb_Mes)
   RemMesAntUTM = 0
   
   '3340329
   For i = 1 To Mes
      
   '3340329
   'Rc = GetRemIVAUTM(Mes, gEmpresa.Ano, RemMesAntUTM)
   Rc = GetRemIVAUTM_New(i, gEmpresa.Ano, RemMesAntUTM, vFmt(Tx_Total))
   '3340329
   'Fecha = DateSerial(gEmpresa.Ano, ItemData(Cb_Mes) + 1, 1)
   If Rc < 0 Then
      RemMesAntUTM = 0
      If Rc = ERR_VALUTM Then
         Lb_UTM.visible = True
      ElseIf Rc = ERR_DEFCUENTA Then
         Lb_DefCta.visible = True
      End If
   Else
      Lb_UTM.visible = False
      Lb_DefCta.visible = False
   End If
   
   Fecha = DateSerial(gEmpresa.Ano, i + 1, 1)   ' se agrega +1 a solicitud de Victor Morales (17 nov. 2011)
   
   'OJO redondeamos a dos decimales
   RemMesAntUTM = vFmt(Format(RemMesAntUTM, DBLFMT2))
   
   
   If GetValMoneda("UTM", ValUTM, Fecha, True) = True Then
      TotRemMesAnt = RemMesAntUTM * ValUTM
      Tx_RemMesAnt = Format(TotRemMesAnt, NEGNUMFMT)
   Else
      Tx_RemMesAnt = 0
      Lb_UTM.visible = True
   End If
   
   Call GetResIVA(i, lAno, TotIVACred, TotIVADeb, TotIEPDGen, TotIEPDTransp)
   
   TotIVACred = TotIVACred - lIVAIrrec          'se resta IVA Irrecuperable que se obtiene en la función LoadValOImp
   Tx_IVACred = Format(TotIVACred, NEGNUMFMT)
   
   TotIVADeb = TotIVADeb - lIVARetParcial - lIVARetTotal        'se resta IVA Retenido Parcial o Total que se obtiene en la función LoadValOImp
   Tx_IVADeb = Format(TotIVADeb, NEGNUMFMT)
   
   Tx_IEPDGenCF = Format(TotIEPDGen, NEGNUMFMT)
   Tx_IEPDTrCF = Format(TotIEPDTransp, NEGNUMFMT)
      
   AjusteIVAMensual = GetAjusteIVAMensual(1)
   Tx_AjusRemCRedFis = Format(AjusteIVAMensual, NUMFMT)
   
'
'   Tx_Total = Format(Abs(TotIVADeb - (TotIVACred + Tx_IEPDGenCF + Tx_IEPDTrCF + TotRemMesAnt + Tx_AjusRemCRedFis)), NEGNUMFMT)
'
'   If TotIVADeb - (TotIVACred + Tx_IEPDGenCF + Tx_IEPDTrCF + TotRemMesAnt) < 0 Then   'remanente
'      Lb_Total = "Remanente periodo siguiente:"
'      Lb_TotalUTM.Visible = True
'      Tx_TotalUTM.Visible = True
'      Lb_RemUTM.Visible = True
'
'      If Lb_UTM.Visible = False Then
'         If ValUTM <> 0 Then
'            TotRemUTM = vFmt(Tx_Total) / ValUTM
'         End If
'         Tx_TotalUTM = Format(TotRemUTM, DBLFMT2)
'      End If
'
'   Else                          'a pagar
'      Lb_Total = "Sub Total a Pagar:"
'      Lb_TotalUTM.Visible = False
'      Tx_TotalUTM.Visible = False
'      Lb_RemUTM.Visible = False
'
'   End If

   '3389677 FPR se creo para que no traspase si es un saldo a pagar al mes siguiente, ya que no es remanente
  '3426794
   'If i = Mes Then
     If Not Mes = 1 And RemIVAAnoAnt = True Then
        If Not remMesAnt Then
            Tx_RemMesAnt.Text = 0
            TotRemMesAnt = 0
        End If
     End If
   'End If
   '3426794
   'FIN 3389677 FPR

   Call Calcular    'FCA   23 sep 2021
   

   
 Next i
 '3340329
   lModVal = False     'FCA 23 sep 2021

End Sub
Private Sub LoadValOImp()
   Dim Where As String
   Dim ResOImp() As ResOImp_t
   Dim i As Integer, j As Integer
   Dim Lib As String, CurrLib As String
   Dim Total As Double
   
   Where = " " & SqlYearLng("FEmision") & " = " & lAno
   
'   If ItemData(Cb_TipoLib) > 0 Then
'
'      If ItemData(Cb_TipoLib) = T_COMPRASVENTAS Then
'         Where = Where & " AND Documento.TipoLib IN (" & LIB_COMPRAS & ", " & LIB_VENTAS & ")"
'      Else
'         Where = Where & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
'      End If
'
'   End If
   
   If ItemData(Cb_Mes) > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & ItemData(Cb_Mes)
   End If
   
'   If lTipoLib > 0 Then
'      Where = Where & " AND Documento.TipoLib = " & lTipoLib
'   End If
   
'   Call GenResOImpEsRecup(Where, ResOImp, False)
   Call GenResOImpEsRecup(Where, ResOImp, True)       'Victor Morales 1 sep 2020
   
   Grid.Redraw = False
   
   Grid.rows = Grid.FixedRows
   CurrLib = ""
   Total = 0
   lIVAIrrec = 0
   lIVARetParcial = 0
   lIVARetTotal = 0
   
   j = Grid.FixedRows
   For i = 0 To UBound(ResOImp)
      
      If ResOImp(i).CodValLib <> 0 Then
               
'        If ResOImp(i).CodValLib <> LIBCOMPRAS_IMPESPDIESEL And ResOImp(i).CodValLib <> LIBCOMPRAS_IMPESPDIESELTRANS Then
            
            Grid.rows = Grid.rows + 1
            
            Lib = ReplaceStr(gTipoLib(ResOImp(i).TipoLib), "Libro de ", "")
            
            If Lib <> CurrLib Then
               CurrLib = Lib
               Grid.TextMatrix(j, C_LIBRO) = Lib
            End If
            
            Grid.TextMatrix(j, C_DESC) = ResOImp(i).DescValLib
            If ResOImp(i).TipoLib = LIB_COMPRAS Then
               Grid.TextMatrix(j, C_RECUPERABLE) = FmtSiNo(Abs(ResOImp(i).EsRecuperable))
            End If
            Grid.TextMatrix(j, C_VALOR) = Format(ResOImp(i).Valor, NEGNUMFMT)
            
            If ResOImp(i).TipoLib = LIB_COMPRAS And (ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC1 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC2 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC3 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC4 Or ResOImp(i).CodValLib = LIBCOMPRAS_IVAIRREC9) Then
               lIVAIrrec = lIVAIrrec + ResOImp(i).Valor
               '628093 se vuelve atras solucion de ado 3429912 ya que afecta calculo de iva credito
               '3429912 ffv
              'If ResOImp(i).EsRecuperable Then
               'lIVAIrrec = lIVAIrrec + ResOImp(i).Valor
              'Else
               'lIVAIrrec = lIVAIrrec
              'End If
              '3429912 ffv
              '628093 fin
            End If
            
            If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_PARCIAL Then
               lIVARetParcial = lIVARetParcial + ResOImp(i).Valor
            End If
            
            If ResOImp(i).TipoLib = LIB_VENTAS And ResOImp(i).TipoIVARetenido = IVARET_TOTAL Then
               lIVARetTotal = lIVARetTotal + ResOImp(i).Valor
            End If
            'gcb21092021
            Total = Total + ResOImp(i).Valor
            '+ vFmt(Tx_AjusRemCRedFis)
            j = j + 1
'         End If
         
      Else
         Exit For
      
      End If
   Next i
   
   'Tx_TotOImp = Format(Total, NEGNUMFMT)
   Call FGrVRows(Grid, 2)
   Grid.Redraw = True
   
End Sub



Private Sub SetUpGrid()
   
   Call FGrSetup(Grid, True)
   Grid.Cols = NCOLS + 1
   
   Grid.ColWidth(C_TIPOLIB) = 0
   Grid.ColWidth(C_CODVALLIB) = 0
   Grid.ColWidth(C_LIBRO) = 1000
   Grid.ColWidth(C_DESC) = 4400
   Grid.ColWidth(C_RECUPERABLE) = 1000
   Grid.ColWidth(C_VALOR) = 1300
   
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_RECUPERABLE) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_LIBRO) = "Libro"
   Grid.TextMatrix(0, C_DESC) = "Impuesto"
   Grid.TextMatrix(0, C_RECUPERABLE) = "Recuperable"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call FGrVRows(Grid, 2)

End Sub

Public Function FView(ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal TipoLib As Integer = 0, Optional ByVal VerBotonRes As Boolean = True)

   lMes = Mes
   lAno = Ano
   lTipoLib = TipoLib
   lVerBotonRes = VerBotonRes
   Me.Show vbModal

End Function

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption)
   End If
End Sub


Private Sub Tx_AjusRemCRedFis_Change()
   lModVal = True          'FCA 23 sep 2021

End Sub

Private Sub Tx_AjusRemCRedFis_KeyPress(KeyAscii As Integer)
'gcb13092021
   Call KeyNum(KeyAscii)

End Sub

Private Sub Tx_AjusRemCRedFis_LostFocus()
   Tx_AjusRemCRedFis = Format(vFmt(Tx_AjusRemCRedFis), NUMFMT)    'FCA 23 sep 2021
   
   Call Calcular
   
End Sub
