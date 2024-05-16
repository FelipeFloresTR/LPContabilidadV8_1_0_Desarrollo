VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Documento"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "FrmDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   240
      Picture         =   "FrmDoc.frx":000C
      ScaleHeight     =   720
      ScaleWidth      =   780
      TabIndex        =   52
      Top             =   420
      Width           =   840
   End
   Begin VB.TextBox Tx_TitMov 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "Detalle del Documento (Compras, Ventas o Retenciones)"
      Top             =   6060
      Width           =   9120
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1815
      Left            =   1335
      TabIndex        =   25
      Top             =   6300
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   10
      Cols            =   6
   End
   Begin VB.CommandButton Bt_Salir 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9345
      TabIndex        =   28
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   9345
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   9345
      TabIndex        =   26
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Fr_Doc 
      Caption         =   "Documento"
      ForeColor       =   &H00FF0000&
      Height          =   3030
      Left            =   1320
      TabIndex        =   32
      Top             =   2880
      Width           =   9120
      Begin VB.ComboBox Cb_Tratamiento 
         Height          =   315
         Left            =   7320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3120
         Width           =   1545
      End
      Begin VB.CommandButton Bt_VerDTE 
         Caption         =   "Ver DTE"
         Height          =   315
         Left            =   7860
         TabIndex        =   24
         ToolTipText     =   "Ver Imagen de DTE si está disponible"
         Top             =   2580
         Width           =   1095
      End
      Begin VB.CheckBox Ch_DTE 
         Height          =   255
         Left            =   6060
         TabIndex        =   13
         Top             =   1200
         Width           =   195
      End
      Begin VB.TextBox Tx_NumDocAsoc 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2580
         Width           =   1275
      End
      Begin VB.ComboBox Cb_TipoDocAsoc 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2580
         Width           =   1335
      End
      Begin VB.CheckBox Ch_DTEDocAsoc 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc. Electrónico"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5580
         TabIndex        =   23
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton Bt_PrtCheque 
         Caption         =   "Imprimir Cheque..."
         Height          =   315
         Left            =   7260
         TabIndex        =   48
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox Tx_CorrInterno 
         Height          =   315
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1170
         Width           =   1095
      End
      Begin VB.CheckBox Ch_DocAnalitico 
         Caption         =   "Incluir en Informe Analítico"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   3195
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   6000
         Picture         =   "FrmDoc.frx":0575
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1620
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2460
         Picture         =   "FrmDoc.frx":087F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1620
         Width           =   230
      End
      Begin VB.ComboBox Cb_Cuentas 
         Height          =   315
         Left            =   4905
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   765
         Width           =   4020
      End
      Begin VB.TextBox Tx_NumDocHasta 
         Height          =   315
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1170
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   7380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   7380
         TabIndex        =   14
         Top             =   1170
         Width           =   1515
      End
      Begin VB.TextBox Tx_Descrip 
         Height          =   450
         Left            =   1365
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   2010
         Width           =   7590
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   750
         Width           =   2835
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1170
         Width           =   1095
      End
      Begin VB.TextBox Tx_FEmision 
         Height          =   315
         Left            =   1365
         TabIndex        =   15
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox Tx_FVenc 
         Height          =   315
         Left            =   4905
         TabIndex        =   17
         Top             =   1620
         Width           =   1095
      End
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   2835
      End
      Begin VB.Label Lbl_Tratamiento 
         Caption         =   "Tratamiento"
         Height          =   255
         Left            =   6120
         TabIndex        =   54
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DTE"
         Height          =   195
         Index           =   12
         Left            =   6300
         TabIndex        =   53
         Top             =   1215
         Width           =   330
      End
      Begin VB.Label Lb_NotaCred 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Asociado:"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N°"
         Height          =   195
         Left            =   2820
         TabIndex        =   50
         Top             =   2640
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Electrónico"
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   49
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Int.:"
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   47
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   11
         Left            =   4320
         TabIndex        =   44
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   10
         Left            =   2580
         TabIndex        =   43
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   8
         Left            =   6780
         TabIndex        =   40
         Top             =   1665
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Index           =   9
         Left            =   6780
         TabIndex        =   39
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   38
         Top             =   2070
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc.:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   37
         Top             =   810
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   36
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión:"
         Height          =   255
         Index           =   4
         Left            =   225
         TabIndex        =   35
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vencimiento:"
         Height          =   195
         Index           =   5
         Left            =   3465
         TabIndex        =   34
         Top             =   1665
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento de:"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Fr_Entidad 
      Caption         =   "Entidad"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   1320
      TabIndex        =   29
      Top             =   360
      Width           =   7815
      Begin VB.TextBox Tx_FechaExport 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Bt_CleanFExport 
         Caption         =   "Volver a exportar"
         Height          =   915
         Left            =   6480
         Picture         =   "FrmDoc.frx":0B89
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Permitir que el sistema vuelva a exportar este documento al año siguiente cuando se realice la apertura"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ComboBox Cb_TipoRelEnt 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1260
         Width           =   1875
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   4380
         TabIndex        =   1
         Top             =   420
         Width           =   195
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   5160
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2595
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         ItemData        =   "FrmDoc.frx":0F38
         Left            =   1320
         List            =   "FrmDoc.frx":0F3A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   810
         Width           =   5055
      End
      Begin VB.CommandButton Bt_NewEnt 
         Caption         =   "Nueva..."
         Height          =   915
         Left            =   6480
         Picture         =   "FrmDoc.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Crear nueva entidad"
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Export:"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   58
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Relación:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   46
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "RUT:"
         Height          =   255
         Left            =   4620
         TabIndex        =   45
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Razón Social:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   30
         Top             =   870
         Width           =   1155
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   1335
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   8100
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   11
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

Const H_SMALL = 4600

Const C_IDMOVDOC = 0
Const C_TIPOVAL = 1
Const C_CODCUENTA = 2
Const C_CUENTA = 3
Const C_DEBE = 4
Const C_HABER = 5

Dim lRc As Integer
Dim lOper As Integer
Dim lcbNombre As ClsCombo
Dim lIdDoc As Long
Dim lMultiplesDocs As Boolean
Dim lTipoLib As Integer
Dim lMes As Integer
Dim lAño As Integer
Dim lValorDoc As Double
Dim lIdEntidad As Long
Dim lClasifEnt As Long

Dim lOldIdCuenta As Long

Dim lPrtGridComp As ClsPrtFlxGrid

Dim lInLoad As Boolean

Dim lUrlDTE As String


Private Sub Bt_Cancel_Click()
   
   If lMultiplesDocs Then
     Call ClearForm
   Else
      lRc = vbCancel
      Unload Me
   End If
End Sub


'3026009
Private Sub Bt_CleanFExport_Click()
Dim Q1 As String
   
   If MsgBox1("¿Esta seguro que desea volver a importar este documento desde el año siguiente cuando haga la apertura?" & vbCrLf & vbCrLf & "Atención: Si este documento ya existe en el año siguiente, y usted genera el comprobante de apertura, éste quedará duplicado." & vbCrLf & vbCrLf & "Si desea volverlo a importar. elimínelo en el año siguiente, antes de generar el comprobante de apertura.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Q1 = "UPDATE Documento SET FExported = 0 WHERE TipoLib = " & LIB_OTROFULL & " AND IdDoc = " & lIdDoc & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   MsgBox1 "El documento podrá ser importado nuevamente desde el año siguiente, al momento de generar el Comprobante de Apertura de ese año.", vbInformation
   
End Sub
'3026009

Private Sub Bt_NewEnt_Click()
   Dim Frm As FrmEntidad
   Dim Row As Integer
   Dim Entidad As Entidad_t
   Dim i As Integer
   Dim Rc As Integer
   Dim AuxRut As String
    
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   
   AuxRut = FmtCID(vFmtCID(Tx_Rut))
   If AuxRut = "0-0" Then
      AuxRut = Trim(Tx_Rut)
   End If
   Entidad.Rut = AuxRut
   
   If Cb_Entidad.ListIndex >= 0 Then
      Entidad.Clasif = ItemData(Cb_Entidad)
   Else
      Entidad.Clasif = SIN_CLASLST
   End If
      
   Rc = Frm.FNew(Entidad, AuxRut)
   If Rc <> vbCancel Then
   
      If Cb_Entidad.ListIndex >= 0 Then
   
         If Entidad.Clasif = ItemData(Cb_Entidad) Then
            
            If Rc = vbOK Then
               Call lcbNombre.AddItem(Entidad.Nombre, Entidad.id, vFmtCID(Entidad.Rut))
               lcbNombre.ListIndex = lcbNombre.NewIndex
            Else
               lcbNombre.SelItem Entidad.id
            End If
            Tx_Rut.Text = Entidad.Rut
            
         Else
            MsgBox1 "La clasificación de la nueva entidad no coincide con la que está seleccionada. Vuelva a seleccionar el tipo de entidad para que la muestre en la lista.", vbOKOnly + vbInformation
         End If
         
      Else
         Cb_Entidad.ListIndex = FindItem(Cb_Entidad, Entidad.Clasif)
         
         Call lcbNombre.AddItem(Entidad.Nombre, Entidad.id, vFmtCID(Entidad.Rut))
         lcbNombre.ListIndex = lcbNombre.NewIndex
         Tx_Rut.Text = Entidad.Rut
      End If
      
   End If
   Set Frm = Nothing
   MousePointer = vbDefault
End Sub


Public Function FEdit(ByVal IdDoc As Long, Optional ByVal TipoLib As Integer = 0) As Integer
   lOper = O_EDIT
   lIdDoc = IdDoc
   lTipoLib = TipoLib
   
   '3026009
   If lTipoLib = LIB_OTROFULL Then
    Bt_CleanFExport.visible = True
   Else
    Bt_CleanFExport.visible = False
   End If
   '3026009
   
   Me.Show vbModal
   
   FEdit = lRc
   
End Function
Public Sub FView(ByVal IdDoc As Long, Optional ByVal TipoLib As Integer = 0)
   
   lOper = O_VIEW
   lIdDoc = IdDoc
   lTipoLib = TipoLib
   Me.Show vbModal
      
End Sub

Public Function FNew(ByVal TipoLib As Integer, IdDoc As Long, Optional ByVal MultiplesDocs As Boolean = True, Optional ByVal Mes As Integer = 0, Optional ByVal Año As Integer = 0, Optional ByVal ValorDoc As Double = 0, Optional ByVal IdEntidad As Long, Optional ByVal TipoLibEnt As Integer = 0) As Integer
   
   lOper = O_NEW
   lTipoLib = TipoLib
   lIdDoc = 0
   lMes = Mes
   lAño = Año
   lMultiplesDocs = MultiplesDocs
   lValorDoc = ValorDoc
   lIdEntidad = IdEntidad
   
   If TipoLibEnt = LIB_COMPRAS Or TipoLibEnt = LIB_RETEN Then
      lClasifEnt = ENT_PROVEEDOR
   ElseIf TipoLibEnt = LIB_VENTAS Then
      lClasifEnt = ENT_CLIENTE
   End If
   
   Me.Show vbModal
   
   IdDoc = lIdDoc
   
   FNew = lRc
   
End Function

Private Sub LoadAll()
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   If lOper = O_NEW Or lIdDoc <= 0 Then
      If lValorDoc <> 0 Then
         Tx_Valor = Format(lValorDoc, NUMFMT)
      End If
      
      If lIdEntidad <> 0 Then
         Call CbSelItem(Cb_Entidad, lClasifEnt)
         Call lcbNombre.SelItem(lIdEntidad)
      End If
      Exit Sub
   End If
   
   Q1 = "SELECT TipoLib, TipoDoc, NumDoc, NumDocHasta, DTE, Documento.IdEntidad, Documento.TipoEntidad, Documento.Total,"
   Q1 = Q1 & " FEmision, FEmisionOri, FVenc, Descrip, Documento.Estado, DocOtrosEnAnalitico, UrlDTE, "
   Q1 = Q1 & " Entidades.Nombre, Entidades.Rut, Entidades.NotValidRut,idCtaBanco, TipoRelEnt, CorrInterno, IdDocAsoc , Tratamiento "
   
   '3125512
   Q1 = Q1 & " ,FExported "
   '3125512
   
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad"
   Q1 = Q1 & " AND Documento.IdEmpresa = Entidades.IdEmpresa "
   Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
   If lTipoLib > 0 Then
    Q1 = Q1 & " AND TipoLib = " & lTipoLib
   End If
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   
   
   'Q2 = Replace(Replace(Q1, "Documento", "DocumentoFull"), ",0", ", DocumentoFull.Tratamiento")
   'Q1 = Q1 & " UNION ALL " & Q2
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Call SelItem(Cb_TipoLib, vFld(Rs("TipoLib")))
      lTipoLib = vFld(Rs("TipoLib"))
      ' se agrega OTROS_Full para boquear campos tipo libro y tipo documento 15719159 nelson ado 3311345 - 3329087
      If vFld(Rs("TipoLib")) = LIB_COMPRAS Or vFld(Rs("TipoLib")) = LIB_VENTAS Or vFld(Rs("TipoLib")) = LIB_RETEN Or vFld(Rs("TipoLib")) = LIB_OTROFULL Then
        '3284709
'        If vFld(Rs("Estado")) = ED_PENDIENTE And vFld(Rs("TipoLib")) = LIB_VENTAS And GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) = TDOC_FAVEXENTA And Year(vFld(Rs("FEmision"))) < gEmpresa.Ano Then
'        Call FrmEnable
'         'Lb_DocAnalitico.Visible = False
'         Ch_DocAnalitico.visible = False
'
'         Call GetDocAsoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")), vFld(Rs("IdDocAsoc")))
'        Else
'
'         lOper = O_VIEW
'         Call FrmEnable
'         'Lb_DocAnalitico.Visible = False
'         Ch_DocAnalitico.visible = False
'
'         Call GetDocAsoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")), vFld(Rs("IdDocAsoc")))
'        End If
        
        '616437
        
        If vFld(Rs("TipoLib")) = LIB_OTROFULL And vFld(Rs("Estado")) = ED_PENDIENTE Then
            Cb_Estado.RemoveItem (ED_PAGADO - 1)
            Cb_Estado.RemoveItem (ED_CENTRALIZADO - 1)
        End If
'
        If vFld(Rs("Estado")) = ED_PAGADO Or vFld(Rs("Estado")) = ED_CENTRALIZADO Then
         lOper = O_VIEW
        End If
        '616437
        
        
        Call FrmEnable
         'Lb_DocAnalitico.Visible = False
         Ch_DocAnalitico.visible = False

         'Call GetDocAsoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")), vFld(Rs("IdDocAsoc")))
         
         If vFld(Rs("TipoLib")) = LIB_OTROFULL Then
            Me.Height = Tx_TitMov.Top + 300
         End If
        
        '3284709
      Else
         Me.Height = Tx_TitMov.Top + 300
         
         i = FindItem(Cb_TipoLib, LIB_COMPRAS)
         If i >= 0 Then
            Cb_TipoLib.RemoveItem (i)
         End If
         i = FindItem(Cb_TipoLib, LIB_VENTAS)
         If i >= 0 Then
            Cb_TipoLib.RemoveItem (i)
         End If
         i = FindItem(Cb_TipoLib, LIB_RETEN)
         If i >= 0 Then
            Cb_TipoLib.RemoveItem (i)
         End If
         
         Ch_DocAnalitico.visible = True
         Ch_DocAnalitico = IIf(vFld(Rs("DocOtrosEnAnalitico")) <> 0, 1, 0)
         
      End If
            
      Call SelItem(Cb_TipoDoc, vFld(Rs("TipoDoc")))
      
      If vFld(Rs("idEntidad")) > 0 Then
         Call SelItem(Cb_Entidad, vFld(Rs("TipoEntidad")))
         Call SelItem(Cb_TipoRelEnt, vFld(Rs("TipoRelEnt")))
      End If
      
      Call lcbNombre.SelItem(vFld(Rs("idEntidad")))
      If vFld(Rs("Rut")) <> "" Then
         Tx_Rut = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0) 'PAM agrego =0 para que dejara bien el manejo de ruts
      End If
      
      Tx_NumDoc = vFld(Rs("NumDoc"))
      Tx_NumDocHasta = vFld(Rs("NumDocHasta"))
      Tx_CorrInterno = vFld(Rs("CorrInterno"))
      
      Ch_DTE = IIf(vFld(Rs("DTE")) <> 0, 1, 0)
      
      If vFld(Rs("TipoLib")) = LIB_COMPRAS Or vFld(Rs("TipoLib")) = LIB_VENTAS Or vFld(Rs("TipoLib")) = LIB_RETEN Then
         Call SetTxDate(Tx_FEmision, vFld(Rs("FEmisionOri")))
      Else
         Call SetTxDate(Tx_FEmision, vFld(Rs("FEmision")))
      End If
      
      If vFld(Rs("FVenc")) <> 0 Then
         Call SetTxDate(Tx_FVenc, vFld(Rs("FVenc")))
      End If
      
      '3125512
      Call SetTxDate(Tx_FechaExport, vFld(Rs("FExported")))
      '3125512
      
      Tx_Valor = Format(vFld(Rs("Total")), NUMFMT)
      
      Tx_Descrip = vFld(Rs("Descrip"), True)
   
      Cb_Estado.ListIndex = FindItem(Cb_Estado, vFld(Rs("Estado")))
      
      If vFld(Rs("idCtaBanco")) > 0 Then
         Call SelItem(Cb_Cuentas, vFld(Rs("idCtaBanco")))
         lOldIdCuenta = vFld(Rs("idCtaBanco"))
      End If
      '***
      
      If vFld(Rs("Tratamiento")) > -1 Then
         Call SelItem(Cb_Tratamiento, vFld(Rs("Tratamiento")))
      End If
   
      lUrlDTE = vFld(Rs("UrlDTE"))
   
   End If
   
   Call CloseRs(Rs)
   
   Call LoadMovs
   
   
End Sub
Private Sub LoadMovs()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   
   Q1 = "SELECT MovDocumento.IdMovDoc, MovDocumento.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion, MovDocumento.Debe, MovDocumento.Haber, IdTipoValLib "
   Q1 = Q1 & " FROM MovDocumento "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = MovDocumento.IdEmpresa AND Cuentas.Ano = MovDocumento.Ano"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   Q1 = Q1 & " WHERE MovDocumento.IdDoc = " & lIdDoc
   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY MovDocumento.Orden "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Rs.EOF = False
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDMOVDOC) = vFld(Rs("IdMovDoc"))
      Grid.TextMatrix(i, C_TIPOVAL) = GetNombreTipoValLib(ItemData(Cb_TipoLib), vFld(Rs("IdTipoValLib")))
      Grid.TextMatrix(i, C_CODCUENTA) = FmtCodCuenta(vFld(Rs("Codigo")))
      Grid.TextMatrix(i, C_CUENTA) = FCase(vFld(Rs("Descripcion"), True))
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), BL_NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), BL_NUMFMT)
      
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call CalcTot
   
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = C_TIPOVAL
   Grid.ColSel = C_TIPOVAL
   
      
End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ConCta As String
   Dim DimDoc As String
   Dim FldArray(3) As AdvTbAddNew_t
   Dim Tabla As String
   
   'If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
    'Tabla = "DocumentoFull"
   'Else
    Tabla = "Documento"
   'End If
   
   If lOper = O_NEW Then

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
            
      'lIdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
          lIdDoc = AdvTbAddNewMult(DbMain, Tabla, "IdDoc", FldArray)
      
   End If
   
   '**PS
   If ItemData(Cb_TipoLib) = LIB_OTROS Or ItemData(Cb_TipoLib) = LIB_OTROFULL And ItemData(Cb_Cuentas) >= 0 Then
      ConCta = ", idCtaBanco =" & ItemData(Cb_Cuentas)
   End If
   '****
   
   Q1 = "UPDATE " & Tabla & " SET "
   Q1 = Q1 & "  TipoDoc =" & ItemData(Cb_TipoDoc)
   Q1 = Q1 & ", TipoLib =" & ItemData(Cb_TipoLib)
   Q1 = Q1 & ", NumDoc ='" & vFmt(Tx_NumDoc) & "'"
   Q1 = Q1 & ", NumDocHasta ='" & vFmt(Tx_NumDocHasta) & "'"
   Q1 = Q1 & ", CorrInterno =" & vFmt(Tx_CorrInterno)
   If ItemData(Cb_Entidad) >= 0 Then
      Q1 = Q1 & ", TipoEntidad =" & ItemData(Cb_Entidad)
      Q1 = Q1 & ", TipoRelEnt =" & ItemData(Cb_TipoRelEnt)
   Else
      Q1 = Q1 & ", TipoRelEnt =0"
   End If
   If lcbNombre.ItemData >= 0 Then
      Q1 = Q1 & ", idEntidad =" & lcbNombre.ItemData
   Else
      Q1 = Q1 & ", idEntidad =0"
   End If
   Q1 = Q1 & ", DocOtrosEnAnalitico =" & IIf(Ch_DocAnalitico <> 0, 1, 0)
   Q1 = Q1 & ", FEmision =" & GetTxDate(Tx_FEmision)
   Q1 = Q1 & ", FEmisionOri =" & GetTxDate(Tx_FEmision)
   Q1 = Q1 & ", FVenc =" & GetTxDate(Tx_FVenc)
   Q1 = Q1 & ", Total =" & vFmt(Tx_Valor)
   Q1 = Q1 & ", Estado =" & ItemData(Cb_Estado)
   '2971346
   'Q1 = Q1 & ", Descrip ='" & ParaSQL(Tx_Descrip) & "'"
   Q1 = Q1 & ", Descrip ='" & Trim(ParaSQL(Tx_Descrip)) & "'"
   '2971346
   Q1 = Q1 & ", SaldoDoc =" & vFmt(Tx_Valor)
   If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
        Q1 = Q1 & ", Tratamiento =" & ItemData(Cb_Tratamiento)
   Else
        Q1 = Q1 & ", Tratamiento = 0"
   End If
   Q1 = Q1 & ConCta
   Q1 = Q1 & " WHERE IdDoc =" & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Tracking 3227543
    Call SeguimientoDocumento(lIdDoc, gEmpresa.id, gEmpresa.Ano, "FrmDoc.SaveAll1", Q1, 1, "", gUsuario.IdUsuario, 1, 2)
    ' fin 3227543
   
         
   'grabamos correaltivo cheque para la cuenta seleccionada, si corresponde
   If CbItemData(Cb_Cuentas) > 0 Then
   
      If CbItemData(Cb_TipoLib) = LIB_OTROS Then
         DimDoc = GetDiminutivoDoc(CbItemData(Cb_TipoLib), CbItemData(Cb_TipoDoc))
         If DimDoc = "CHE" Or DimDoc = "CHF" Then
   
            
            'eliminamos el correlativo de la cuenta anterior si no es new y cambió la cuenta del banco
            If lOper <> O_NEW Then
               If CbItemData(Cb_Cuentas) <> lOldIdCuenta Then
                  
                  Q1 = "UPDATE Cuentas SET CorrelativoCheque = 0 WHERE IdCuenta = " & lOldIdCuenta
                  Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                  Call ExecSQL(DbMain, Q1)
                  
               End If
            
            Else
               Q1 = "UPDATE Cuentas SET CorrelativoCheque = " & Val(Tx_NumDoc) & " WHERE IdCuenta = " & CbItemData(Cb_Cuentas)
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
               Call ExecSQL(DbMain, Q1)
            
            End If
         End If
      End If
      
   End If
            
      
End Sub
Private Sub LoadParametrosODF()
Dim Q1 As String
Dim Rs As Recordset
   If lTipoLib = LIB_OTROFULL Then
       Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='INFANAODF' "
       Set Rs = OpenRs(DbMain, Q1)
          
       If Rs.EOF = False Then
            Ch_DocAnalitico = vFld(Rs("Valor"))
       End If
       Call CloseRs(Rs)
    
       Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='TRATAMODF' "
       Set Rs = OpenRs(DbMain, Q1)
          
       If Rs.EOF = False Then
            Call SelItem(Cb_Tratamiento, vFld(Rs("Valor")))
       End If
       Call CloseRs(Rs)
   End If

End Sub
Private Sub ControlesXTipoLib(lTipoLib As Integer)

If lTipoLib = LIB_OTROFULL Then
    Bt_PrtCheque.visible = False
    Bt_VerDTE.visible = False
    Lbl_Tratamiento.Left = 6300
    Lbl_Tratamiento.Top = 2640
    Cb_Tratamiento.Left = 7380
    Cb_Tratamiento.Top = 2580
    Label1(10).visible = False
    Tx_NumDocHasta.visible = False
    Me.Cb_TipoDoc.ListIndex = 0
    Cb_TipoDocAsoc.Enabled = False
Else
    Bt_PrtCheque.visible = True
    Bt_VerDTE.visible = True
    Lbl_Tratamiento.Left = 6120
    Lbl_Tratamiento.Top = 3120
    Cb_Tratamiento.Left = 7320
    Cb_Tratamiento.Top = 3120
    Label1(10).visible = True
    Tx_NumDocHasta.visible = True
    Cb_TipoDocAsoc.Enabled = True
End If

End Sub
Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   Call lcbNombre.AddItem(" ", -1)
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, NotValidRut FROM Entidades "
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY Nombre "
      Call lcbNombre.FillCombo(DbMain, Q1, -1)
   End If
End Sub
Private Function valida() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdxTipoDoc As Integer
   Dim Fecha As Long
   
   valida = False
   
   If Not ValidaIngresoDoc() Then
      Exit Function
   End If
               
   If Cb_TipoDoc.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un tipo de documento.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If

   If Tx_Rut <> "" And ItemData(Cb_Nombre) <= 0 Then
      MsgBox1 "El RUT ingresado no tiene una entidad asociada.", vbExclamation
      Tx_Rut.SetFocus
      Exit Function
   End If
      
   'Feña
   If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
   
      If Cb_Entidad.Text = "" Then
        MsgBox1 "Favor Ingresar una Entidad.", vbExclamation
        Cb_Entidad.SetFocus
        Exit Function
      End If
      
      If Tx_Rut = "" Then
        MsgBox1 "Favor Ingresar un Rut.", vbExclamation
        Tx_Rut.SetFocus
        Exit Function
      End If
      
      If ItemData(Cb_Nombre) <= 0 Then
        MsgBox1 "Favor Ingresar una Razon Social.", vbExclamation
        Cb_Nombre.SetFocus
        Exit Function
      End If
      
      If ItemData(Cb_TipoRelEnt) <= 0 Then
        MsgBox1 "Favor Ingresar un Tipo Relación.", vbExclamation
        Cb_TipoRelEnt.SetFocus
        Exit Function
      End If
   End If

   IdxTipoDoc = GetTipoDoc(lTipoLib, ItemData(Cb_TipoDoc))
   
   If IdxTipoDoc >= 0 Then
   
      If gTipoDoc(IdxTipoDoc).ExigeRUT And Tx_Rut = "" Then
         MsgBox1 "Debe ingresar una entidad.", vbExclamation
         Tx_Rut.SetFocus
         Exit Function
      End If
      
   End If
   
   If Val(Tx_NumDoc) = 0 Then
      MsgBox1 "Debe ingresar un número de documento.", vbExclamation
      Tx_NumDoc.SetFocus
      Exit Function
   End If
   
   Fecha = GetTxDate(Tx_FEmision)
   
   If Year(Fecha) <> gEmpresa.Ano Then
      If MsgBox1("ATENCIÓN:" & vbNewLine & vbNewLine & "La fecha de emisión del documento no corresponde al año en que está trabajando." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
         Tx_FEmision.SetFocus
         Exit Function
      End If
   End If
   
   If Trim(Tx_Valor) = "" Then
      MsgBox1 "Falta ingresar el valor del documento.", vbExclamation
      Tx_Valor.SetFocus
      Exit Function
   Else
        If lTipoLib = LIB_OTROFULL Then
            If Val(Tx_Valor) <= 0 Then
                    MsgBox1 "El Valor No puede ser menor o igual a 0", vbExclamation
                    Tx_Valor.SetFocus
                    Exit Function
'            ElseIf CInt(Val(Tx_Valor)) <> Val(Tx_Valor) Then
'                MsgBox1 "El Valor No puede ser decimal", vbExclamation
'                Tx_Valor.SetFocus
'                Exit Function
            End If
        End If
   End If
   
   'veamos si este documento ya ha sido ingresado
   If lOper = O_NEW Or lOper = O_EDIT Then
   
      Q1 = "SELECT IdDoc FROM Documento "
      Q1 = Q1 & " WHERE TipoLib=" & ItemData(Cb_TipoLib)
      Q1 = Q1 & " AND TipoDoc=" & ItemData(Cb_TipoDoc)
      Q1 = Q1 & " AND NumDoc='" & Trim(Tx_NumDoc) & "'"
      Q1 = Q1 & " AND IdDoc <> " & lIdDoc       'para que no sea el mismo, en caso de edit
      If CbItemData(Cb_Cuentas) > 0 Then
         Q1 = Q1 & " AND IdCtaBanco =" & CbItemData(Cb_Cuentas)
      Else
         Q1 = Q1 & " AND IdCtaBanco =0"
         If lcbNombre.ItemData >= 0 Then
            Q1 = Q1 & " AND IdEntidad =" & lcbNombre.ItemData
         Else
            Q1 = Q1 & " AND IdEntidad =0"
         End If
      End If
      
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then   'ya existe
            
         If CbItemData(Cb_Cuentas) = 0 Then
            If lcbNombre.ItemData < 0 Then
               If MsgBox1("Este documento ya ha sido ingresado al sistema, sin una entidad o una cuenta de banco asociada. Es posible que esté duplicado." & vbNewLine & vbNewLine & "¿Desea verificar los datos antes de grabar?", vbQuestion + vbYesNo) = vbYes Then
                  Call CloseRs(Rs)
                  Exit Function
               End If
            Else
               MsgBox1 "Este documento ya ha sido ingresado al sistema.", vbExclamation + vbOKOnly
               Call CloseRs(Rs)
               Exit Function
            End If
         Else
            MsgBox1 "Este documento ya ha sido ingresado al sistema.", vbExclamation + vbOKOnly
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If Trim(Tx_FEmision) = "" Then
      MsgBox1 "Debe ingresar fecha de emisión.", vbExclamation
      Tx_FEmision.SetFocus
      Exit Function
   End If
   
'   If Month(Tx_FEmision) <> GetMesActual Then
'      Call MsgBox1("La fecha del documento no correpsonde al mes actualmente abierto.", vbOKOnly + vbExclamation)
'      Exit Function
'   End If
   
   'If Trim(Tx_FVenc) = "" Then
   '   MsgBox1 "Debe ingresar fecha de vecimiento, ya que usted seleccionó un tipo de documento.", vbExclamation
   '   Tx_FVenc.SetFocus
   '   Exit Function
   'End If
   
   If GetTxDate(Tx_FVenc) > 0 And GetTxDate(Tx_FEmision) > GetTxDate(Tx_FVenc) Then
      MsgBox1 "Fecha de emisión mayor a la fecha de vencimiento.", vbExclamation
      Tx_FVenc.SetFocus
      Exit Function
   End If
      
   valida = True
   
End Function
Private Sub FrmEnable()
   Dim bool As Integer
   
   bool = ((lOper = O_EDIT And ItemData(Cb_Estado) <> ED_ANULADO) Or lOper = O_NEW)
   
   If lOper = O_EDIT Or lOper = O_NEW Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         bool = False
      End If
   End If
      
   Cb_TipoLib.Enabled = bool
   Cb_TipoDoc.Enabled = bool
   
   'bool = bool And (ItemData(Cb_TipoDoc) <> -1)
     
   Call SetTxRO(Tx_Rut, Not bool)
   
   Ch_Rut.Enabled = bool
   
   Ch_DocAnalitico.Enabled = bool
   
   Call SetTxRO(Tx_NumDoc, Not bool)
   Call SetTxRO(Tx_FEmision, Not bool)
   Call SetTxRO(Tx_FVenc, Not bool)
   Call SetTxRO(Tx_NumDocHasta, Not bool)
   Call SetTxRO(Tx_Valor, Not bool)
   Call SetTxRO(Tx_Descrip, Not bool)
   'ffv 2855046
   Call SetTxRO(Tx_CorrInterno, Not bool)
   'ffv 2855046
   
   Ch_DTE.Enabled = bool
   
   Cb_Entidad.Enabled = bool
   Cb_Nombre.Enabled = bool
   Me.Cb_Tratamiento.Enabled = bool
   
   If CbItemData(Cb_TipoLib) = LIB_OTROS Or CbItemData(Cb_TipoLib) = LIB_REMU Then
      Cb_Estado.Enabled = True
   Else
      Cb_Estado.Enabled = bool
   End If
   
   Cb_Cuentas.Enabled = bool
   
   Bt_Fecha(0).Enabled = bool
   Bt_Fecha(1).Enabled = bool
   Bt_NewEnt.Enabled = bool
   
   
   If bool = False And Cb_Estado.Enabled = False Then
      Bt_OK.visible = False
      Bt_Cancel.Caption = "Cerrar"
      Bt_Cancel.Top = Bt_OK.Top
   End If
   
End Sub

Private Sub Bt_Salir_Click()

   If (Val(Tx_NumDoc) <> 0 Or vFmt(Tx_Valor) <> 0 Or Trim(Tx_Descrip) <> "") And lOper = O_NEW Then
      If MsgBox1("¿Desea guardar el documento actual?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
         If valida() Then
            Call SaveAll
            lRc = vbOK
            Unload Me
         End If
      Else
         lRc = vbCancel
         Unload Me
      End If
   Else
      lRc = vbCancel
      Unload Me
   End If

End Sub

Private Sub Bt_VerDTE_Click()

      If lUrlDTE = "" Then
         MsgBox1 "No es posible ver el PDF de este documento.", vbExclamation
         Exit Sub
      End If
            
      Call ShellExecute(Me.hWnd, "open", lUrlDTE, "", "", 1)
      
      Me.MousePointer = vbDefault

End Sub

Private Sub Cb_Cuentas_Click()
   Dim DimDoc As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim CorrelatCheque As Long
   
   If CbItemData(Cb_Cuentas) > 0 Then
   
      If CbItemData(Cb_TipoLib) = LIB_OTROS Then
         DimDoc = GetDiminutivoDoc(CbItemData(Cb_TipoLib), CbItemData(Cb_TipoDoc))
         If DimDoc = "CHE" Or DimDoc = "CHF" Then
            
            Q1 = "SELECT CorrelativoCheque FROM Cuentas WHERE IdCuenta=" & CbItemData(Cb_Cuentas)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               CorrelatCheque = vFld(Rs("CorrelativoCheque"))
            End If
            Call CloseRs(Rs)
            
            If CorrelatCheque > 0 Then
            
               If lOper = O_NEW Then
                  Tx_NumDoc = CorrelatCheque + 1
               End If
               
               If Not lInLoad Then
                  MsgBox1 "ATENCIÓN: Verifique el correlativo del cheque propuesto por el sistema.", vbExclamation + vbOKOnly
               End If
            ElseIf lOper = O_NEW Then
               Tx_NumDoc = ""
            End If
            
         End If
      End If
   End If
         

End Sub

Private Sub Cb_Entidad_Click()
      
   If Cb_Entidad.ListIndex >= 0 Then
      Call SelCbEntidad(ItemData(Cb_Entidad))
   Else
      Cb_Nombre.Clear
   End If
   
End Sub
Private Sub cb_Nombre_Click()
   
   Tx_Rut = ""
   
   If lcbNombre.ListIndex >= 0 Then
      If lcbNombre.Matrix(M_RUT) <> "" Then
         Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
         Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
      End If
   End If
   
End Sub
Private Sub Cb_TipoLib_Click()
   Dim Q1 As String
   Dim i As Integer
   Dim Tipo As String
   Dim TipoLib As Integer
   Dim tipLib As Long
   Cb_TipoDoc.Clear
   
   TipoLib = ItemData(Cb_TipoLib)
   
   If TipoLib > 0 Then
   
      Call FillTipoDoc(Cb_TipoDoc, TipoLib, True, False)
      Cb_TipoDoc.ListIndex = -1
      
      'PS **
      Cb_Cuentas.Enabled = TipoLib = LIB_OTROS
      
        If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
          tipLib = 8
        End If
        Call FillCtasConcil(Cb_Cuentas, 2, tipLib)
       
        If tipLib = LIB_OTROFULL And lOper = O_NEW Then
         Cb_Cuentas.ListIndex = FindItem(Cb_Cuentas, CuentaOdfNew(IIf(ItemData(Cb_Tratamiento) = 1, "CTAODFACTI", "CTAODFPASI")))
        End If
      
      '**
   End If
   
   Call FrmEnable
   Call ControlesXTipoLib(TipoLib)
   
   If TipoLib = LIB_OTROFULL Then
        Ch_DocAnalitico = 1
        Cb_Tratamiento.ListIndex = 0
        Cb_Estado.ListIndex = 1
   End If
   
End Sub


Private Sub Cb_TipoDoc_Click()
   Call FrmEnable
   
   If InStr(LCase(Cb_TipoDoc), "cheque") > 0 Then
      Bt_PrtCheque.Enabled = True
   Else
      Bt_PrtCheque.Enabled = True
   End If
   
End Sub



Private Sub Cb_Tratamiento_Click()
   If ItemData(Cb_TipoLib) = LIB_OTROFULL And lOper = O_NEW Then
    Cb_Cuentas.ListIndex = FindItem(Cb_Cuentas, CuentaOdfNew(IIf(ItemData(Cb_Tratamiento) = 1, "CTAODFACTI", "CTAODFPASI")))
   End If
End Sub

Private Sub Ch_Rut_Click()
  ' lcbNombre.ListIndex = -1
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim tipLib As Long
   
   lInLoad = True
   tipLib = 0
      
   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))
   
   Bt_PrtCheque.visible = gFunciones.PrtCheque
   
   'SE LLENA COMBOS
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Cb_Entidad.AddItem ""
   Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = -1
   For i = ENT_CLIENTE To ENT_OTRO
      Cb_Entidad.AddItem gClasifEnt(i)
      Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = i
   Next i
   Cb_Entidad.ListIndex = -1     'para no seleccionar ninguno al partir
   
   Cb_TipoRelEnt.AddItem ""
   Cb_TipoRelEnt.ItemData(Cb_TipoRelEnt.NewIndex) = 0
   For i = TRE_EMISOR To TRE_OTRO
      Cb_TipoRelEnt.AddItem gTipoRelEnt(i)
      Cb_TipoRelEnt.ItemData(Cb_TipoRelEnt.NewIndex) = i
   Next i
   Cb_TipoRelEnt.ListIndex = -1     'para no seleccionar ninguno al partir
   
   Ch_Rut = 1
 
   For i = 1 To UBound(gTipoLibNew)
      If lOper <> O_NEW Or (lOper = O_NEW And i <> LIB_COMPRAS And i <> LIB_VENTAS And i <> LIB_RETEN) Then
         Cb_TipoLib.AddItem ReplaceStr(gTipoLibNew(i).Nombre, "Libro de ", "")
         Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).id 'i
      End If
   Next i
   
   For i = 1 To UBound(gTratamiento)
      Cb_Tratamiento.AddItem ReplaceStr(gTratamiento(i).Nombre, "Libro de ", "")
      Cb_Tratamiento.ItemData(Cb_Tratamiento.NewIndex) = gTratamiento(i).id 'i
   Next i
   Cb_Tratamiento.ListIndex = 1
   
      For i = 1 To MAX_ESTADODOC
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   Cb_Estado.ListIndex = FindItem(Cb_Estado, ED_PENDIENTE)
   
   Cb_TipoLib.ListIndex = -1
   If lTipoLib > 0 Then
      For i = 0 To Cb_TipoLib.ListCount - 1
         If Cb_TipoLib.ItemData(i) = lTipoLib Then
            Cb_TipoLib.ListIndex = i
            Exit For
         End If
      Next i
      
   End If
   
'   For i = 1 To MAX_ESTADODOC
'      Cb_Estado.AddItem gEstadoDoc(i)
'      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
'   Next i
'   Cb_Estado.ListIndex = FindItem(Cb_Estado, ED_PENDIENTE)
   
         
   Bt_Salir.visible = False
               
   Select Case lOper
      Case O_NEW
         Caption = "Nuevo Documento"
         
         If lTipoLib = LIB_OTROFULL Then
            Ch_DocAnalitico = 1
            Cb_Tratamiento.ListIndex = 0
            '3026009
            Bt_CleanFExport.visible = False
            '3026009
            
            '2855046
           ' Cb_Estado.ListIndex = 2
            'Cb_Estado.ListIndex = 1
            '2855046
         End If
         If lMultiplesDocs Then
            Bt_OK.Caption = "Agregar"
            Bt_Salir.visible = True
         End If
         
         Me.Height = Tx_TitMov.Top + 300
         
      Case O_EDIT
         Caption = "Editar Documento"
         Bt_OK.Caption = "Aceptar"
      Case Else
         Caption = "Ver Documento"
         
   End Select
           
   If lMes > 0 And lAño > 0 And lAño * 100# + lMes <> Format(Now, "yyyymm") Then
      Call SetTxDate(Tx_FEmision, DateSerial(lAño, lMes, 1))
   Else
      Call SetTxDate(Tx_FEmision, Now)
   End If
   
   Call SetUpGrid
   
   'PS
   If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
     tipLib = 8
   End If
   Call FillCtasConcil(Cb_Cuentas, 2, tipLib)
   Cb_Cuentas.Enabled = ItemData(Cb_TipoLib) = LIB_OTROS
   
   If tipLib = LIB_OTROFULL And lOper = O_NEW Then
    Cb_Cuentas.ListIndex = FindItem(Cb_Cuentas, CuentaOdfNew(IIf(ItemData(Cb_Tratamiento) = 1, "CTAODFACTI", "CTAODFPASI")))
   End If
   '***
   
   Call ParametrosODF
   
   Call LoadAll
   
   Call FrmEnable
   
   Call ControlesXTipoLib(lTipoLib)
   lInLoad = False
   
'    If lOper = O_EDIT Then
'    '3284709
'        If ItemData(Cb_Estado) = ED_PENDIENTE And CbItemData(Cb_TipoLib) = LIB_VENTAS And GetDiminutivoDoc(CbItemData(Cb_TipoLib), CbItemData(Cb_TipoDoc)) = TDOC_FAVEXENTA And Year(GetTxDate(Tx_FEmision)) < gEmpresa.Ano Then
'         Cb_Estado.Enabled = True
'         MsgBox1 "Documento proviene del año anterior, favor de dejar estado del documento en centralizado ya que se encuentra en estado pendiente. Favor de ejecutar Recalcular Saldo.", vbInformation
'            Call SelItem(Cb_Estado, ED_CENTRALIZADO)
'        Else
'         'Cb_Estado.Enabled = False
'        End If
'        '3284709
'   End If
   
End Sub

Private Function CuentaOdfNew(Tipo As String) As Long
Dim Q1 As String
Dim Rs As Recordset
   Q1 = "Select Valor From Paramempresa Where Tipo = '" & Tipo & "' "
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
       CuentaOdfNew = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)


End Function

Private Sub ParametrosODF()
Dim Q1 As String
Dim Rs As Recordset

   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='INFANAODF' "
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
       Ch_DocAnalitico = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='ESTADOODF' "
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
        Call SelItem(Cb_Estado, vFld(Rs("Valor")))
   End If
   Call CloseRs(Rs)


End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_FEmision)
   Else
      Call Frm.TxSelDate(Tx_FVenc)
   End If
   
   Set Frm = Nothing
End Sub
Private Sub Bt_OK_Click()

   If ItemData(Cb_TipoLib) = LIB_OTROFULL And ItemData(Cb_Cuentas) = 0 Then
      If MsgBox1("No Ingreso una Cuenta para este documento. Desea Continuar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
        Exit Sub
      End If
   End If
      
   If Not valida() Then
      Exit Sub
   End If
      
   Call SaveAll
   
   If lMultiplesDocs Then
      Call ClearForm
   Else
      lRc = vbOK
      Unload Me
   End If
   
End Sub




Private Sub Tx_CorrInterno_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)

End Sub

Private Sub Tx_FEmision_GotFocus()
   Call DtGotFocus(Tx_FEmision)
End Sub

Private Sub Tx_FEmision_LostFocus()
   
   If Trim$(Tx_FEmision) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FEmision)
   
End Sub

Private Sub Tx_FEmision_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FVenc_GotFocus()
   Call DtGotFocus(Tx_FVenc)
End Sub

Private Sub Tx_FVenc_LostFocus()
   
   If Trim$(Tx_FVenc) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FVenc)
   
End Sub

Private Sub Tx_FVenc_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_NumDoc_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub


Private Sub Tx_NumDocHasta_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)

End Sub

Private Sub Tx_Rut_Change()
  '    lcbNombre.ListIndex = -1

End Sub

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Exit Sub
   End If
   
'   If Not MsgValidCID(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
'
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut, Ch_Rut <> 0) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEnt = 0
   
   If Rs.EOF = False Then   'existe
      IdEnt = vFld(Rs("IdEntidad"))
            
      'seleccionamos el tipo de entidad y esto llena la lista de nombres de entidades
      For i = 1 To MAX_ENTCLASIF   'el cero tiene blanco
         If vFld(Rs("Clasif" & Cb_Entidad.ItemData(i))) <> 0 Then
            Cb_Entidad.ListIndex = i
            Exit For
         End If
      Next i
   
      'ahora seleccionamos la entidad
      For i = 0 To Cb_Nombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
   ElseIf MsgBox1("Este RUT no ha sido ingresado al sistema." & vbNewLine & vbNewLine & "¿Desea crear una nueva entidad?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
      Call Bt_NewEnt_Click
   Else
      Cb_Entidad.ListIndex = -1
   End If
      
      
   Call CloseRs(Rs)
   
   If Ch_Rut <> 0 Then
      AuxRut = FmtCID(vFmtCID(Tx_Rut))
      If AuxRut <> "0-0" Then
         Tx_Rut = AuxRut
      End If
   End If
   
End Sub

Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_Rut_LostFocus
      KeyAscii = 0
   ElseIf Ch_Rut <> 0 Then
      Call KeyCID(KeyAscii)
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
   
End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_Valor_LostFocus()
   Tx_Valor = Format(vFmt(Tx_Valor), NUMFMT)
   
End Sub

Private Function ClearForm()

   Tx_NumDoc = ""
   Tx_Valor = ""
   Call SetTxDate(Tx_FEmision, Now)
   Call SetTxDate(Tx_FVenc, 0)
   Tx_Descrip = ""

End Function
Private Sub SetUpGrid()
   Dim i As Integer
   Dim WCodCuenta As Integer
   Dim WCuenta As Integer
    
   Grid.ColWidth(C_IDMOVDOC) = 0
   
   WCodCuenta = Me.TextWidth(gFmtCodigoCta) + 300
   WCuenta = 1450
   
   Grid.ColWidth(C_TIPOVAL) = 1600
   Grid.ColWidth(C_CODCUENTA) = WCodCuenta + 200
   Grid.ColWidth(C_CUENTA) = WCuenta * 2 + 200
   Grid.ColWidth(C_DEBE) = 1400
   Grid.ColWidth(C_HABER) = 1400
         
   Grid.ColAlignment(C_TIPOVAL) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_TIPOVAL) = "Tipo Valor"
   Grid.TextMatrix(0, C_CODCUENTA) = "Cod.Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
 
   Call FGrVRows(Grid)
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
      
   GridTot.TextMatrix(0, C_TIPOVAL) = "TOTAL"

End Sub

Private Sub CalcTot()
   Dim i As Integer
   Dim TotDebe As Double
   Dim TotHaber As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_IDMOVDOC) = "" Then
         Exit For
      End If
      
      TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
      TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
   Next i
   
   GridTot.TextMatrix(0, C_DEBE) = Format(TotDebe, NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(TotHaber, NUMFMT)

End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

Private Sub Bt_PrtCheque_Click()
   Dim Frm As FrmPrtCheque
   Dim NumCheque As Long
   Dim Ref As String
   
   If Bt_OK.visible Then
      If MsgBox1("Se grabarán los datos antes de imprimir el cheque." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      If Not valida() Then
         Exit Sub
      End If
         
      Call SaveAll
   End If
   
   NumCheque = Val(Tx_NumDoc)
   Ref = ""
   
   Set Frm = New FrmPrtCheque
   Call Frm.FPrint(False, Nothing, NumCheque, GetTxDate(Tx_FEmision), Cb_Nombre, Ref, Cb_Cuentas, vFmt(Tx_Valor), 0)
   Set Frm = Nothing
   
End Sub

Private Function GetDocAsoc(ByVal TipoLib As Integer, ByVal TipoDoc As Integer, ByVal IdDoc As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim DimDoc As String
   Dim TipoDocAsoc As Integer
   
   If IdDoc <= 0 Then
      Exit Function
   End If
   
   DimDoc = GetDiminutivoDoc(TipoLib, TipoDoc)
   If (DimDoc = "NCC" Or DimDoc = "NDC") Then
      Call CbAddItem(Cb_TipoDocAsoc, " ", 0)
      TipoDocAsoc = FindTipoDoc(TipoLib, "FAC")
      Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
      TipoDocAsoc = FindTipoDoc(TipoLib, "FCE")
      Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
   
   '3030542
   'ElseIf DimDoc = "NCV" Or DimDoc = "NDV" Then
    ElseIf DimDoc = "NCV" Or DimDoc = "NDV" Or DimDoc = "NCE" Or DimDoc = "NDE" Then
   '3030542
      Call CbAddItem(Cb_TipoDocAsoc, " ", 0)
      TipoDocAsoc = FindTipoDoc(TipoLib, "FAV")
      Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
      TipoDocAsoc = FindTipoDoc(TipoLib, "FVE")
      Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
      '3030542
      TipoDocAsoc = FindTipoDoc(TipoLib, "EXP")
      Call CbAddItem(Cb_TipoDocAsoc, GetNombreTipoDoc(TipoLib, TipoDocAsoc), TipoDocAsoc)
      '3030542
   End If

   
   Cb_TipoDocAsoc.ListIndex = -1
   Tx_NumDocAsoc = ""
   Ch_DTEDocAsoc = 0
   
   Q1 = "SELECT TipoDoc, NumDoc, Estado, DTE FROM Documento WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If Cb_TipoDocAsoc.ListCount > 0 Then
         For i = 0 To Cb_TipoDocAsoc.ListCount - 1
            If vFld(Rs("TipoDoc")) = Cb_TipoDocAsoc.ItemData(i) Then
               Cb_TipoDocAsoc.ListIndex = i
               Tx_NumDocAsoc = vFld(Rs("NumDoc"))
               Ch_DTEDocAsoc = IIf(vFld(Rs("DTE")) <> 0, 1, 0)
               Exit For
            End If
         Next i
      Else
         MsgBox1 "Documento asociado inválido.", vbExclamation
      End If
      
   Else
      MsgBox1 "Documento asociado inválido.", vbExclamation
   End If
   
   Call CloseRs(Rs)

End Function

