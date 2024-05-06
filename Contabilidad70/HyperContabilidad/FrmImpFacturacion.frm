VERSION 5.00
Begin VB.Form FrmImpFacturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Documentos desde Facturación"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Archivo a Importar"
      Height          =   2415
      Left            =   360
      TabIndex        =   12
      Top             =   1740
      Width           =   9315
      Begin VB.CommandButton Bt_BrowseDB 
         Height          =   435
         Left            =   7920
         Picture         =   "FrmImpFacturacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ubique la carpeta Datos del Sistema de Facturación y seleccione el archivo TRFactura.mdb"
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton Op_ImpDesde 
         Caption         =   "Conectar directo a Base de Datos de Facturación ubicada en:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox Tx_DbFactura 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   8895
      End
      Begin VB.OptionButton Op_ImpDesde 
         Caption         =   "Importar desde Archivo exportado en TRFacturación (debe haberlo traspasado a esta carpeta):"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Debe almacenar el archivo exportado en TRFacturación en la carpeta que se indica en este campo"
         Top             =   1440
         Width           =   7275
      End
      Begin VB.TextBox Tx_FName 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1740
         Width           =   8895
      End
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8100
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   375
      Left            =   8100
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   480
      Width           =   4215
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   420
         Width           =   795
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   14
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   360
      Picture         =   "FrmImpFacturacion.frx":056A
      ScaleHeight     =   660
      ScaleWidth      =   660
      TabIndex        =   9
      Top             =   600
      Width           =   660
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Index           =   0
      Left            =   1260
      TabIndex        =   8
      Top             =   480
      Width           =   2055
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Docs.  de Compras"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Docs. de Ventas"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmImpFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const IMP_DB = 1
Const IMP_FILE = 2
 
 
Const EDTE_EMITIDO = 3   'estado de un DTE en Facturación


Dim lRc As Integer
Dim lMsgAdv As String

Dim lIdCtaAfecto As Long
Dim lIdCtaExento As Long
Dim lIdCtaTotal As Long
Dim lDbName As String
Dim lTipoLib As Integer

Dim lInLoad As Boolean
Dim lTxDbFactura As String



Private Sub Bt_Cancelar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim i As Integer
   Dim Rc As Integer
   
   If lIdCtaAfecto = 0 Or lIdCtaExento = 0 Or lIdCtaTotal = 0 Then
      MsgBox1 "Falta definir las cuentas por omisión para los libros de compras y ventas (Afecto, Exento y Total)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
      Exit Sub
   End If
   
   If lTipoLib > 0 Then
   
      If lMsgAdv = False Then    'este mensaje se muestra sólo una vez
      
         If MsgBox1("Para realizar la importación del " & gTipoLib(lTipoLib) & ", nadie más debe estar trabajando en este libro para esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      
      End If

      lMsgAdv = True
      
      Me.MousePointer = vbHourglass
      
      Rc = ImpFactMes()
      
      Me.MousePointer = vbDefault
      
      If Rc Then
         If Op_ImpDesde(IMP_FILE) = True Then
            gDbFacturacion = ""
         Else
            gDbFacturacion = Tx_DbFactura
         End If
         Call SetIniString(gIniFile, "Config", "PathFactura", gDbFacturacion)
      End If

            
   End If

End Sub


Private Sub Cb_Mes_Click()
   Call GenDbName
   
End Sub

Private Sub Form_Activate()
   
   If MsgBox1("ATENCIÓN: Sólo se importarán los DTE de Facturación que se encuentren en estado EMITIDO." & vbCrLf & vbCrLf & "Asegúrese de haber revisado y actualizado el estado de cada DTE en TRFacturación, ingresando a DTE Emitidos >> Ver Detalle Estado" & vbCrLf & "antes de realizar esta operación. " & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim Q1 As String
   Dim i As Integer
   
   lInLoad = True
   lMsgAdv = False
      
   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   
   Tx_Ano = gEmpresa.Ano
   
   DbMainDate = GetDbNow(DbMain)
   
   lTipoLib = LIB_VENTAS
   Op_Libros(lTipoLib) = 1
   
   lInLoad = False
   
   If gDbFacturacion <> "" Then
      Op_ImpDesde(IMP_DB) = True
      Tx_DbFactura = gDbFacturacion
   Else
      Op_ImpDesde(IMP_FILE) = True
   End If
   
   Call GenDbName
   
'   Call AddItem(Cb_Sucursal, " ", 0)
'   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales ORDER BY Descripcion"
'   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)

End Sub

Private Function ImpFactMes() As Boolean
   Dim Db As Database
   Dim Q1 As String
   Dim Rs As Recordset
   Dim StrMes As String
   Dim ImpEnable As Boolean
   Dim RsDoc As Recordset
   Dim IdEnt As Long
   Dim Nombre As String
   Dim NotValidRut As Boolean
   Dim i As Integer
   Dim RsNewDoc As Recordset
   Dim IdDoc As Long
   Dim j As Integer
   Dim RsDao As dao.Recordset
   Dim FldName As String
   Dim NUpd As Long
   Dim NIns As Long
   Dim ImpDbPath As String
   Dim StrAño As String, Ano As Integer
   Dim ChkSumCtasLoc As Long
   Dim ChkSumCtasImp As Long
   Dim Rc As Integer
   Dim Rut As String
   Dim Mes As Integer
   Dim Where As String
   Dim TblDTE As String, TblDetDTE As String, TblEnt As String, TblEntOld As String, TblEmpresas As String
   Dim ConnStr As String
   Dim TipoDocFCC As Integer, TipoDocFCV As Integer
   Dim IdEntidad As Long, IdEmpresaDTE As Long
   Dim IdCtaAfectoEntidad As Long, IdCtaExentoEntidad As Long, IdCtaTotalEntidad As Long, IdPropIVAEntidad As Integer, EsDelGiroEntidad As Boolean
   Dim IdCtaAfecto As Long, IdCtaExento As Long, IdCtaTotal As Long, IdPropIVA As Integer, EsDelGiro As Boolean
   Dim IdAreaNegAfectoEntidad As Long, IdAreaNegExentoEntidad As Long, IdAreaNegTotalEntidad As Long, IdCCostoAfectoEntidad As Long, IdCCostoExentoEntidad As Long, IdCCostoTotalEntidad  As Long
   Dim DocEdit As Boolean
   Dim NombreEntidad As String, Descrip As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim TmpTblDoc As String, TmpTblMovDoc As String
   Dim FldArray(4) As AdvTbAddNew_t
   
   Dim sMsg As String
  On Error GoTo Error_Handler
   
   ImpFactMes = False
   
   If lDbName = "" Then
      Exit Function
   End If
   
   ConnStr = gEmpresa.ConnStr

   
   Rut = gEmpresa.Rut
   Ano = Val(Tx_Ano)
   Mes = ItemData(Cb_Mes)
   StrAño = Ano
   StrMes = StrAño & Right("0" & Mes, 2)

   If Op_ImpDesde(IMP_FILE) Then
      If lTipoLib = LIB_VENTAS Then
         TblDTE = "DTE_" & StrMes
         TblDetDTE = "DetDTE_" & StrMes
         TblEnt = "Entidades_" & StrMes
         TblEmpresas = "EmpresaDTESel"
      Else
         TblDTE = "DTERecibidos_" & StrMes
         TblDetDTE = ""
         TblEnt = "Entidades_" & StrMes
         TblEmpresas = "EmpresaDTESel"
      End If
   Else
      If lTipoLib = LIB_VENTAS Then
         TblDTE = "DTE"
         TblDetDTE = "DetDTE"
         TblEnt = "Entidades"
         TblEmpresas = "Empresa"
      Else
         TblDTE = "DTERecibidos"
         TblDetDTE = ""
         TblEnt = "Entidades"
         TblEmpresas = "Empresa"
      End If
   End If
      
   If Not ExistFile(lDbName) Then
      MsgBox1 "No se encontró el archivo:" & vbNewLine & vbNewLine & "        " & lDbName & vbNewLine & vbNewLine & "Verifique que el archivo se encuentre en la carpeta especificada y vuelva a intentarlo.", vbExclamation + vbOKOnly
      Exit Function
   End If

   ImpEnable = LockAction(DbMain, LK_IMPLIBROS, Mes)
   
   If ImpEnable = False Then    'alguien más está importando este mes
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_EXPLIBROS, Mes) & "'. No se realizará la importación.", vbInformation
      Exit Function
   End If
   
   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & lDbName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
   
   'la base de datos de exportación no tiene password

   Set Db = OpenDatabase(lDbName, False, False, ConnStr)

   'Obtenemos el Id de la Empresa en el sistema de facturación
   Q1 = "SELECT Id FROM " & TblEmpresas & " WHERE Rut = '" & gEmpresa.Rut & "'"
   Set RsDao = OpenRsDao(Db, Q1)
   If Not RsDao.EOF Then
      IdEmpresaDTE = vFldDao(RsDao("Id"))
   End If
   Call CloseRs(RsDao)
   
'   If W.InDesign Then
'      IdEmpresaDTE = 2
'   End If
   
   If IdEmpresaDTE = 0 Then
      MsgBox1 "Error al obtener identificación de la empresa en facturación.", vbExclamation
      Exit Function
   End If
   
     
   Q1 = "SELECT * FROM " & TblEnt
   Q1 = Q1 & " WHERE " & TblEnt & ".IdEmpresa = " & IdEmpresaDTE

   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
      
      Q1 = "SELECT * FROM Entidades"
      Q1 = Q1 & " WHERE Entidades.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " AND Entidades.Rut = '" & vFldDao(RsDao("Rut")) & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then   'ya existe, la actualizamos
      
         Q1 = "UPDATE Entidades SET "
         Q1 = Q1 & " Entidades.Codigo = '" & vFldDao(RsDao("Codigo")) & "'"
         Q1 = Q1 & ", Entidades.Nombre = '" & vFldDao(RsDao("Nombre")) & "'"
         Q1 = Q1 & ", Entidades.Direccion = '" & vFldDao(RsDao("Direccion")) & "'"
         Q1 = Q1 & ", Entidades.Region = " & vFldDao(RsDao("Region"))
         Q1 = Q1 & ", Entidades.Comuna = " & vFldDao(RsDao("Comuna"))
         Q1 = Q1 & ", Entidades.Ciudad = '" & vFldDao(RsDao("Ciudad")) & "'"
         Q1 = Q1 & ", Entidades.Telefonos = '" & vFldDao(RsDao("Telefonos")) & "'"
         Q1 = Q1 & ", Entidades.Fax = '" & vFldDao(RsDao("Fax")) & "'"
         Q1 = Q1 & ", Entidades.ActEcon = " & vFldDao(RsDao("ActEcon"))
         Q1 = Q1 & ", Entidades.CodActEcon = '" & vFldDao(RsDao("CodActEcon")) & "'"
         Q1 = Q1 & ", Entidades.DomPostal = '" & vFldDao(RsDao("DomPostal")) & "'"
         Q1 = Q1 & ", Entidades.ComPostal = " & vFldDao(RsDao("ComPostal"))
         Q1 = Q1 & ", Entidades.Email = '" & vFldDao(RsDao("Email")) & "'"
         Q1 = Q1 & ", Entidades.Web = '" & vFldDao(RsDao("Web")) & "'"
         Q1 = Q1 & ", Entidades.Giro = '" & vFldDao(RsDao("Giro")) & "'"
         Q1 = Q1 & ", Entidades.EsSupermercado = " & vFldDao(RsDao("EsSupermercado"))
         Q1 = Q1 & " WHERE Entidades.IdEmpresa = " & gEmpresa.id
         Q1 = Q1 & " AND Entidades.Rut = '" & vFldDao(RsDao("Rut")) & "'"
         
         Call ExecSQL(DbMain, Q1, False)
      
      Else    'entidad no existe, la agregamos
      
         ERR.Clear
         
         Q1 = "INSERT INTO Entidades "
         Q1 = Q1 & "(IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut, EsSupermercado, EntRelacionada)"
         Q1 = Q1 & " VALUES (" & gEmpresa.id & ",'" & vFldDao(RsDao("Rut")) & "','" & vFldDao(RsDao("Codigo")) & "','" & ParaSQL(vFldDao(RsDao("Nombre"))) & "','" & ParaSQL(vFldDao(RsDao("Direccion"))) & "'," & vFldDao(RsDao("Region"))
         Q1 = Q1 & "," & vFldDao(RsDao("Comuna")) & ",'" & vFldDao(RsDao("Ciudad")) & "','" & vFldDao(RsDao("Telefonos")) & "','" & vFldDao(RsDao("Fax")) & "'," & vFldDao(RsDao("ActEcon")) & ",'" & vFldDao(RsDao("CodActEcon")) & "'"
         Q1 = Q1 & ",'" & vFldDao(RsDao("DomPostal")) & "'," & vFldDao(RsDao("ComPostal")) & ",'" & vFldDao(RsDao("Email")) & "','" & vFldDao(RsDao("Web")) & "'," & vFldDao(RsDao("Estado")) & ",'" & vFldDao(RsDao("Obs")) & "'"
         Q1 = Q1 & "," & vFldDao(RsDao("Clasif0")) & "," & vFldDao(RsDao("Clasif1")) & "," & vFldDao(RsDao("Clasif2")) & "," & vFldDao(RsDao("Clasif3")) & "," & vFldDao(RsDao("Clasif4")) & "," & vFldDao(RsDao("Clasif5"))
         Q1 = Q1 & ",'" & vFldDao(RsDao("Giro")) & "'," & Abs(vFldDao(RsDao("NotValidRut"))) & "," & Abs(vFldDao(RsDao("EsSupermercado"))) & "," & Abs(vFldDao(RsDao("EntRelacionada"))) & ")"
      
         Call ExecSQL(DbMain, Q1, False)
               
      End If
      
      Call CloseRs(Rs)
   
      RsDao.MoveNext
      
   Loop
   
   Call CloseRs(RsDao)
   
'   If lTipoLib = LIB_VENTAS Then   esto es si la base de Facturación fuera SQLServer
'      'OJO Estado debe ser = a EMITIDO en versión FINAL
'      Where = SqlYearLng(TblDTE & ".Fecha") & " = " & Ano & " AND " & SqlMonthLng(TblDTE & ".Fecha") & " = " & Mes & " AND " & TblDTE & ".IdEmpresa = " & IdEmpresaDTE
'      Where = Where & " AND " & TblDTE & ".IdEstado = " & EDTE_EMITIDO
'   Else
'      Where = SqlYearLng(TblDTE & ".FEmision") & " = " & Ano & " AND " & SqlMonthLng(TblDTE & ".FEmision") & " = " & Mes & " AND " & TblDTE & ".IdEmpresa = " & IdEmpresaDTE
'   End If
   
   If lTipoLib = LIB_VENTAS Then
      Where = " Year(" & TblDTE & ".Fecha) = " & Ano & " AND Month(" & TblDTE & ".Fecha) = " & Mes & " AND " & TblDTE & ".IdEmpresa = " & IdEmpresaDTE
      Where = Where & " AND " & TblDTE & ".IdEstado = " & EDTE_EMITIDO
   Else
      Where = " Year(" & TblDTE & ".FEmision) = " & Ano & " AND Month(" & TblDTE & ".FEmision) = " & Mes & " AND " & TblDTE & ".IdEmpresa = " & IdEmpresaDTE
   End If
   
   
   'Insertamos todos los DTEs en una tabla de Documentos temporal
   
   'primero creamos la tabla vacía
   
   TmpTblDoc = DbGenTmpName2(SQL_ACCESS, "tmpdoc_")
   TmpTblMovDoc = DbGenTmpName2(SQL_ACCESS, "tmpmovdoc_")
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblDoc)
   
   Q1 = "SELECT Documento.* INTO " & TmpTblDoc & " FROM Documento WHERE 1=0 "
   Call ExecSQL(DbMain, Q1)
  
   
   'insertamos los documentos nuevos en la tabla temporal
   Q1 = "  SELECT IdDTE, TipoLib, TipoDoc, Folio"
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", Fecha"
      Q1 = Q1 & ", RUT "
   Else
      Q1 = Q1 & ", FEmision "
      Q1 = Q1 & ", RUTEmisor"
   End If
   
   Q1 = Q1 & ", IdEntidad"
   Q1 = Q1 & ", Exento"
   Q1 = Q1 & ", Neto"
   Q1 = Q1 & ", IVA"
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", ImpAdicional"
   Else
      Q1 = Q1 & ", Impuestos"
   End If
   
   Q1 = Q1 & ", Total"
   Q1 = Q1 & ", UrlDTE "
   
   Q1 = Q1 & " FROM " & TblDTE
   
   Q1 = Q1 & " WHERE " & Where & " AND TipoLib = " & lTipoLib
   Q1 = Q1 & " AND Total <> 0 "
   
   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
   
#If DATACON = 2 Then
      Q1 = "SET IDENTITY_INSERT " & TmpTblDoc & " ON;"
#Else
      Q1 = ""
#End If
   
      Q1 = Q1 & " INSERT INTO " & TmpTblDoc
      Q1 = Q1 & "(IdDoc, IdEmpresa, Ano, TipoLib, TipoDoc, NumDoc"
      Q1 = Q1 & ", FEmision, FEmisionOri "
      Q1 = Q1 & ", RutEntidad "
      Q1 = Q1 & ", IdEntidad, Estado "
      Q1 = Q1 & ", Exento, IdCuentaExento"
      Q1 = Q1 & ", Afecto, IdCuentaAfecto "
      Q1 = Q1 & ", IVA, IdCuentaIVA "
      Q1 = Q1 & ", OtroImp, IdCuentaOtroImp "
      Q1 = Q1 & ", Total, IdCuentaTotal "
      Q1 = Q1 & ", IdUsuario, FechaCreacion"
      Q1 = Q1 & ", UrlDTE "
      Q1 = Q1 & ", FImpFacturacion, DTE, MovEdited )"
   
      Q1 = Q1 & " VALUES(" & vFldDao(RsDao("IdDTE")) & "," & gEmpresa.id & "," & gEmpresa.Ano & "," & vFldDao(RsDao("TipoLib")) & "," & vFldDao(RsDao("TipoDoc")) & ",'" & vFldDao(RsDao("Folio")) & "'"
   
      If lTipoLib = LIB_VENTAS Then
         Q1 = Q1 & ", " & vFldDao(RsDao("Fecha")) & "," & vFldDao(RsDao("Fecha"))
         Q1 = Q1 & ", " & vFldDao(RsDao("RUT"))
      Else
         Q1 = Q1 & ", " & vFldDao(RsDao("FEmision")) & "," & vFldDao(RsDao("FEmision"))
         Q1 = Q1 & ", '" & vFldDao(RsDao("RUTEmisor")) & "'"
      End If
      
      Q1 = Q1 & ", " & vFldDao(RsDao("IdEntidad")) & ", " & ED_PENDIENTE
      Q1 = Q1 & ", " & vFldDao(RsDao("Exento")) & ", " & lIdCtaExento
      Q1 = Q1 & ", " & vFldDao(RsDao("Neto")) & ", " & lIdCtaAfecto
      Q1 = Q1 & ", " & vFldDao(RsDao("IVA")) & ", " & gCtasBas.IdCtaIVADeb
      
      If lTipoLib = LIB_VENTAS Then
         Q1 = Q1 & ", " & vFldDao(RsDao("ImpAdicional")) & ", " & gCtasBas.IdCtaOtrosImpDeb
      Else
         Q1 = Q1 & ", " & vFldDao(RsDao("Impuestos")) & ", " & gCtasBas.IdCtaOtrosImpCred
      End If
      
      Q1 = Q1 & ", " & vFldDao(RsDao("Total")) & ", " & lIdCtaTotal
      Q1 = Q1 & ", " & gUsuario.IdUsuario & ", " & Int(DbMainDate)
      Q1 = Q1 & ", '" & ParaSQL(vFldDao(RsDao("UrlDTE"))) & "'"
      Q1 = Q1 & ", " & Int(DbMainDate) & ", 1, 0);"
      
#If DATACON = 2 Then
      Q1 = Q1 & "SET IDENTITY_INSERT " & TmpTblDoc & " OFF;"
#End If
      
      Call ExecSQL(DbMain, Q1)
      
      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)

   
'   'actualizamos IdEntidad en los documentos por si hay nuevos y IdEntidad de los DTE no calza con los de Contabilidad
   Tbl = TmpTblDoc
   sFrom = TmpTblDoc & " INNER JOIN Entidades ON " & TmpTblDoc & ".RutEntidad = Entidades.Rut "
   sSet = TmpTblDoc & ".IdEntidad = Entidades.IdEntidad"
   sWhere = " WHERE Entidades.IdEmpresa = " & gEmpresa.id

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'eliminamos las cuentas de exento, IVA y OtroImp si el valor del campo es 0
   Q1 = "UPDATE " & TmpTblDoc
   Q1 = Q1 & " SET IdCuentaExento = 0 WHERE Exento = 0 OR Exento IS NULL"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE " & TmpTblDoc
   Q1 = Q1 & " SET IdCuentaIVA = 0 WHERE IVA = 0 OR IVA IS NULL"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE " & TmpTblDoc
   Q1 = Q1 & " SET IdCuentaOtroImp = 0 WHERE OtroImp = 0 OR OtroImp IS NULL"
   Call ExecSQL(DbMain, Q1)
   
   'si trajimos Facturas de Compra desde Facturación (libro de Ventas), debemos combiarlas por FCC de libro de Compras
   If lTipoLib = LIB_VENTAS Then
      TipoDocFCC = FindTipoDoc(LIB_COMPRAS, "FCC")
      TipoDocFCV = FindTipoDoc(LIB_VENTAS, "FCV")
   
      Q1 = "UPDATE " & TmpTblDoc
      Q1 = Q1 & " SET TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & ", TipoDoc = " & TipoDocFCC
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND TipoDoc = " & TipoDocFCV
      
      Call ExecSQL(DbMain, Q1)
   End If
   
   
    
   'Ahora agregamos los movimientos en la tabla MovDocumento
   
   'primero creamos la tabla vacía
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblMovDoc)
   
   Q1 = "SELECT 0 as EnProceso, MovDocumento.*  INTO " & TmpTblMovDoc & " FROM MovDocumento WHERE 1=0 "
   Call ExecSQL(DbMain, Q1)
   
   'Primero insertamos todos los EXENTOS de los documentos de Compra o Venta
   Q1 = "SELECT " & TblDTE & ".IdDTE as IdDoc, 1 as Orden, " & lIdCtaExento & " as IdCuenta "
   Q1 = Q1 & ", ' ' as Glosa "
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', " & TblDTE & ".Exento, 0) as Debe "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', 0, " & TblDTE & ".Exento) as Haber "
      Q1 = Q1 & ", " & LIBVENTAS_EXENTO & " as IdTipoValLib "
   
   Else 'LIB_COMPRAS
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, " & TblDTE & ".Exento, 0) as Haber "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, 0, " & TblDTE & ".Exento) as Debe "
      Q1 = Q1 & ", " & LIBCOMPRAS_EXENTO & " as IdTipoValLib "
      
   End If
      
   Q1 = Q1 & " FROM " & TblDTE
   Q1 = Q1 & " INNER JOIN TipoDocs ON " & TblDTE & ".TipoLib = TipoDocs.TipoLib AND " & TblDTE & ".TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE " & TblDTE & ".TipoLib = " & lTipoLib & " AND (" & TblDTE & ".Exento <> 0 AND NOT " & TblDTE & ".Exento IS NULL) "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " ORDER BY " & TblDTE & ".IdDTE "
   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
   
      Q1 = "INSERT INTO " & TmpTblMovDoc
      Q1 = Q1 & "( IdDoc, Orden, IdCuenta "
      Q1 = Q1 & ", Glosa "
      Q1 = Q1 & ", Debe "
      Q1 = Q1 & ", Haber "
      Q1 = Q1 & ", IdTipoValLib, EnProceso )"
      Q1 = Q1 & " VALUES("
      Q1 = Q1 & vFldDao(RsDao("IdDoc")) & "," & vFldDao(RsDao("Orden")) & "," & vFldDao(RsDao("IdCuenta"))
      Q1 = Q1 & ",'" & ParaSQL(vFldDao(RsDao("Glosa"))) & "'"
      Q1 = Q1 & "," & vFldDao(RsDao("Debe"))
      Q1 = Q1 & "," & vFldDao(RsDao("Haber"))
      Q1 = Q1 & "," & vFldDao(RsDao("IdTipoValLib")) & ", 0)"
      
      Call ExecSQL(DbMain, Q1)
      
      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)
  
   'Segundo, insertamos todos los AFECTOS de los documentos de Compra o Venta
   Q1 = "  SELECT " & TblDTE & ".IdDTE as IdDoc, 2 as Orden, " & lIdCtaAfecto & " as IdCuenta "
   Q1 = Q1 & ", ' ' as Glosa "
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', " & TblDTE & ".Neto, 0) as Debe "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', 0, " & TblDTE & ".Neto) as Haber "
      Q1 = Q1 & ", " & LIBVENTAS_AFECTO & " as IdTipoValLib "
   
   Else 'LIB_COMPRAS
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, " & TblDTE & ".Neto, 0) as Haber "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, 0, " & TblDTE & ".Neto) as Debe "
      Q1 = Q1 & ", " & LIBCOMPRAS_AFECTO & " as IdTipoValLib "
      
   End If
      
   Q1 = Q1 & " FROM " & TblDTE
   Q1 = Q1 & " INNER JOIN TipoDocs ON " & TblDTE & ".TipoLib = TipoDocs.TipoLib AND " & TblDTE & ".TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE " & TblDTE & ".TipoLib = " & lTipoLib & " AND " & TblDTE & ".Neto <> 0 "
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " ORDER BY " & TblDTE & ".IdDTE "
   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
   
      Q1 = "INSERT INTO " & TmpTblMovDoc
      Q1 = Q1 & "( IdDoc, Orden, IdCuenta "
      Q1 = Q1 & ", Glosa "
      Q1 = Q1 & ", Debe "
      Q1 = Q1 & ", Haber "
      Q1 = Q1 & ", IdTipoValLib, EnProceso )"
      Q1 = Q1 & " VALUES("
      Q1 = Q1 & vFldDao(RsDao("IdDoc")) & "," & vFldDao(RsDao("Orden")) & "," & vFldDao(RsDao("IdCuenta"))
      Q1 = Q1 & ",'" & ParaSQL(vFldDao(RsDao("Glosa"))) & "'"
      Q1 = Q1 & "," & vFldDao(RsDao("Debe"))
      Q1 = Q1 & "," & vFldDao(RsDao("Haber"))
      Q1 = Q1 & "," & vFldDao(RsDao("IdTipoValLib")) & ", 0)"
      
      Call ExecSQL(DbMain, Q1)
      
      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)
     
   'Tercero, insertamos el IVA de cada documento de Compra o Venta
   Q1 = "  SELECT " & TblDTE & ".IdDTE as IdDoc, 3 as Orden "
   Q1 = Q1 & ", ' ' as Glosa "
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", " & gCtasBas.IdCtaIVADeb & " as IdCuenta "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', " & TblDTE & ".IVA, 0) as Debe "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', 0, " & TblDTE & ".IVA) as Haber "
      Q1 = Q1 & ", " & LIBVENTAS_IVADEBFISC & " as IdTipoValLib "
   
   Else 'LIB_COMPRAS
      Q1 = Q1 & ", " & gCtasBas.IdCtaIVADeb & " as IdCuenta "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, " & TblDTE & ".IVA, 0) as Haber "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, 0, " & TblDTE & ".IVA) as Debe "
      Q1 = Q1 & ", " & LIBCOMPRAS_IVACREDFISC & " as IdTipoValLib "
      
   End If
      
   Q1 = Q1 & " FROM " & TblDTE
   Q1 = Q1 & " INNER JOIN TipoDocs ON " & TblDTE & ".TipoLib = TipoDocs.TipoLib AND " & TblDTE & ".TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE " & TblDTE & ".TipoLib = " & lTipoLib
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " ORDER BY " & TblDTE & ".IdDTE "
   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
   
      Q1 = "INSERT INTO " & TmpTblMovDoc
      Q1 = Q1 & "( IdDoc, Orden, IdCuenta "
      Q1 = Q1 & ", Glosa "
      Q1 = Q1 & ", Debe "
      Q1 = Q1 & ", Haber "
      Q1 = Q1 & ", IdTipoValLib, ENProceso )"
      Q1 = Q1 & " VALUES("
      Q1 = Q1 & vFldDao(RsDao("IdDoc")) & "," & vFldDao(RsDao("Orden")) & "," & vFldDao(RsDao("IdCuenta"))
      Q1 = Q1 & ",'" & ParaSQL(vFldDao(RsDao("Glosa"))) & "'"
      Q1 = Q1 & "," & vFldDao(RsDao("Debe"))
      Q1 = Q1 & "," & vFldDao(RsDao("Haber"))
      Q1 = Q1 & "," & vFldDao(RsDao("IdTipoValLib")) & ",0 )"
      
      Call ExecSQL(DbMain, Q1)
      
      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)
  
   'Cuarto, insertamos todos los OTROS IMPUESTOS (o IMPUESTOS ADICIONALES) de los documentos de Compra o Venta
   If lTipoLib = LIB_VENTAS Then
      Q1 = "  SELECT " & TblDetDTE & ".IdDTE as IdDoc, 4 as Orden"
      'Q1 = Q1 & ", Count(*) & ' Producto(s)' as Glosa "
      Q1 = Q1 & ", ' ' as Glosa "
      
'      If lTipoLib = LIB_VENTAS Then
         Q1 = Q1 & ", " & gCtasBas.IdCtaOtrosImpDeb & " as IdCuenta "
         Q1 = Q1 & ", Sum(iif(TipoDocs.EsRebaja AND TipoDocs.Diminutivo <> 'FCV', " & TblDetDTE & ".MontoImpAdic, 0)) as Debe "
         Q1 = Q1 & ", Sum(iif(TipoDocs.EsRebaja AND TipoDocs.Diminutivo <> 'FCV', 0, " & TblDetDTE & ".MontoImpAdic)) as Haber "
      
'      Else 'LIB_COMPRAS   'este detalle de otros impuestos aún no lo tenemos en compras
'         Q1 = Q1 & ", " & gCtasBas.IdCtaOtrosImpCred & " as IdCuenta "
'         Q1 = Q1 & ", Sum(iif(TipoDocs.EsRebaja, " & TblDetDTE & ".MontoImpAdic, 0)) as Haber "
'         Q1 = Q1 & ", Sum(iif(TipoDocs.EsRebaja, 0, " & TblDetDTE & ".MontoImpAdic)) as Debe "
'      End If
         
      Q1 = Q1 & ", " & TblDetDTE & ".IdImpAdic as IdTipoValLib "
      Q1 = Q1 & " FROM (" & TblDetDTE & " INNER JOIN " & TblDTE & " ON " & TblDetDTE & ".IdDTE = " & TblDTE & ".IdDTE)"
      Q1 = Q1 & " INNER JOIN TipoDocs ON " & TblDTE & ".TipoLib = TipoDocs.TipoLib AND " & TblDTE & ".TipoDoc = TipoDocs.TipoDoc "
      Q1 = Q1 & " WHERE " & TblDTE & ".TipoLib = " & lTipoLib & " AND NOT (IdImpAdic IS NULL OR IdImpAdic = 0)"
      Q1 = Q1 & " AND " & Where
      Q1 = Q1 & " GROUP BY " & TblDetDTE & ".IdDTE, " & TblDetDTE & ".IdImpAdic "
      Q1 = Q1 & " ORDER BY " & TblDetDTE & ".IdDTE, " & TblDetDTE & ".IdImpAdic "
      Set RsDao = OpenRsDao(Db, Q1)
      
      Do While Not RsDao.EOF
      
         Q1 = "INSERT INTO " & TmpTblMovDoc
         Q1 = Q1 & "( IdDoc, Orden, IdCuenta "
         Q1 = Q1 & ", Glosa "
         Q1 = Q1 & ", Debe "
         Q1 = Q1 & ", Haber "
         Q1 = Q1 & ", IdTipoValLib, EnProceso )"
         Q1 = Q1 & " VALUES("
         Q1 = Q1 & vFldDao(RsDao("IdDoc")) & "," & vFldDao(RsDao("Orden")) & "," & vFldDao(RsDao("IdCuenta"))
         Q1 = Q1 & ",'" & ParaSQL(vFldDao(RsDao("Glosa"))) & "'"
         Q1 = Q1 & "," & vFldDao(RsDao("Debe"))
         Q1 = Q1 & "," & vFldDao(RsDao("Haber"))
         Q1 = Q1 & "," & vFldDao(RsDao("IdTipoValLib")) & ", 0 )"
         
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
      Loop
      
      Call CloseRs(RsDao)
      
   End If
  
   'Quinto, insertamos todos los TOTALES de los documentos de Compra o Venta
   Q1 = "  SELECT " & TblDTE & ".IdDTE as IdDoc, 5 as Orden, " & lIdCtaTotal & " as IdCuenta "
   Q1 = Q1 & ", ' ' as Glosa, -1 as EsTotalDoc "
   
   If lTipoLib = LIB_VENTAS Then
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', " & TblDTE & ".Total, 0) as Haber "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja OR TipoDocs.Diminutivo = 'FCV', 0, " & TblDTE & ".Total) as Debe "
      Q1 = Q1 & ", " & LIBVENTAS_TOTAL & " as IdTipoValLib "
   
   Else 'LIB_COMPRAS
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, " & TblDTE & ".Total, 0) as Debe "
      Q1 = Q1 & ", iif(TipoDocs.EsRebaja, 0, " & TblDTE & ".Total) as Haber "
      Q1 = Q1 & ", " & LIBCOMPRAS_TOTAL & " as IdTipoValLib "
      
   End If
      
   Q1 = Q1 & " FROM " & TblDTE
   Q1 = Q1 & " INNER JOIN TipoDocs ON " & TblDTE & ".TipoLib = TipoDocs.TipoLib AND " & TblDTE & ".TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE " & TblDTE & ".TipoLib = " & lTipoLib
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " ORDER BY " & TblDTE & ".IdDTE "
   Set RsDao = OpenRsDao(Db, Q1)
   
   Do While Not RsDao.EOF
   
      Q1 = "INSERT INTO " & TmpTblMovDoc
      Q1 = Q1 & "( IdDoc, Orden, IdCuenta "
      Q1 = Q1 & ", Glosa "
      Q1 = Q1 & ", Debe "
      Q1 = Q1 & ", Haber "
      Q1 = Q1 & ", EsTotalDoc "
      Q1 = Q1 & ", IdTipoValLib, EnProceso )"
      Q1 = Q1 & " VALUES("
      Q1 = Q1 & vFldDao(RsDao("IdDoc")) & "," & vFldDao(RsDao("Orden")) & "," & vFldDao(RsDao("IdCuenta"))
      Q1 = Q1 & ",'" & ParaSQL(vFldDao(RsDao("Glosa"))) & "'"
      Q1 = Q1 & "," & vFldDao(RsDao("Debe"))
      Q1 = Q1 & "," & vFldDao(RsDao("Haber"))
      Q1 = Q1 & "," & vFldDao(RsDao("EsTotalDoc"))
      Q1 = Q1 & "," & vFldDao(RsDao("IdTipoValLib")) & ", 0 )"
      
      Call ExecSQL(DbMain, Q1)
      
      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)
   
   Call CloseDb(Db)
   
   
'   ahora importamos uno a uno
   Q1 = "SELECT * FROM " & TmpTblDoc
   Set Rs = OpenRs(DbMain, Q1)

   NUpd = 0
   NIns = 0

   Do While Not Rs.EOF

      Q1 = "SELECT * "
      Q1 = Q1 & " FROM Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True)
      Q1 = Q1 & " WHERE TipoLib = " & vFld(Rs("TipoLib"))
      Q1 = Q1 & " AND TipoDoc = " & vFld(Rs("TipoDoc"))
      Q1 = Q1 & " AND NumDoc = '" & vFld(Rs("NumDoc")) & "'"

'      If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
'         Q1 = Q1 & " AND NumDocHasta = '" & vFld(Rs("NumDocHasta")) & "'"
'      Else
'         Q1 = Q1 & " AND (NumDocHasta = '0' OR NumDocHasta IS NULL) "
'      End If

'      If vFld(Rs("IdEntidad")) <> 0 Then
         Q1 = Q1 & " AND Entidades.Rut = '" & vFld(Rs("RutEntidad")) & "'"
'      Else
'         Q1 = Q1 & " AND (Documento.IdEntidad IS NULL OR Documento.IdEntidad = 0)"
'      End If

      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " AND Documento.Ano = " & gEmpresa.Ano


      'obtenemos IdDoc
      Set RsDoc = OpenRs(DbMain, Q1)

      IdDoc = 0

      If RsDoc.EOF Then   'no está, lo agregamos
      
         'insertamos el Doc

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
         
         FldArray(4).FldName = "Estado"       'Fecha recepción utilizamos la fecha de Acuse del cliente  (no se usa la FechaRec)
         FldArray(4).FldValue = ED_PENDIENTE
         FldArray(4).FldIsNum = True
                              
         IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
            
         NIns = NIns + 1

         Call CloseRs(RsDoc)

         'agregamos el resto de los campos, incluyendo cuentas, áreas de negocio, centros de costo, PropIVA y EsDelGiro por omisión por proveedor si las hay

         If IdDoc <> 0 Then
   
            'armamos el UPDATE campo a campo
            
            DocEdit = False
            
            Q1 = ""
   
            For i = 0 To Rs.Fields.Count - 1
               FldName = UCase(Rs.Fields(i).Name)
               If FldName = "RUT" Then   'inicio de campos adicionales de exportación
                  Exit For
               End If
               If FldName <> "IDDOC" And FldName <> "ESTADO" And FldName <> "FECHACREACION" And FldName <> "IDUSUARIO" And FldName <> "FIMPFACTURACION" Then
                  If FldIsString(Rs.Fields(i)) Then
                     Q1 = Q1 & "," & Rs.Fields(i).Name & "= '" & ParaSQL(vFld(Rs(Rs.Fields(i).Name))) & "'"
                  Else
                     Q1 = Q1 & "," & Rs.Fields(i).Name & "= " & vFld(Rs(Rs.Fields(i).Name))
                  End If
               End If
            Next i
   
            Q1 = Q1 & ", FImpFacturacion=" & CLng(Int(Now))   'guardamos la fecha de la última actualización
   
            Q1 = "UPDATE Documento SET " & Mid(Q1, 2)   'sacamos la primera coma
            Q1 = Q1 & " WHERE IdDoc = " & IdDoc
            Call ExecSQL(DbMain, Q1)
         
            'ahora actualizamos las cuentas del documento de acuerdo a la entidad, si corresponde
            'vemos si hay cuentas para entidad para marcar como editado el documento, de tal manera que no se genere automáticamente los movimientos y estos datos se pierdan
            IdEntidad = vFld(Rs("IdEntidad"))
            Call GetCuentasEntidad(lTipoLib, IdEntidad, IdCtaAfectoEntidad, IdCtaExentoEntidad, IdCtaTotalEntidad, IdPropIVAEntidad, EsDelGiroEntidad, NombreEntidad)
            
            If IdCtaAfectoEntidad <> 0 Then
               IdCtaAfecto = IdCtaAfectoEntidad
               DocEdit = True
            Else
               IdCtaAfecto = lIdCtaAfecto
            End If
      
            If IdCtaExentoEntidad <> 0 Then
               IdCtaExento = IdCtaExentoEntidad
               DocEdit = True
            Else
               IdCtaExento = lIdCtaExento
            End If
      
            If IdCtaTotalEntidad <> 0 Then
               IdCtaTotal = IdCtaTotalEntidad
               DocEdit = True
            Else
               IdCtaTotal = lIdCtaTotal
            End If
      
            If IdPropIVAEntidad <> 0 Then
               IdPropIVA = IdPropIVAEntidad
            Else
               IdPropIVA = 0
            End If
      
            If EsDelGiroEntidad <> 0 Then
               EsDelGiro = EsDelGiroEntidad
            Else
               EsDelGiro = 0
            End If
      
            
            'ahora obtenemos area de negocio y centro de costo del documento de acuerdo a la entidad, si corresponde
            Call GetANegCCostoEntidad(lTipoLib, IdEntidad, IdAreaNegAfectoEntidad, IdAreaNegExentoEntidad, IdAreaNegTotalEntidad, IdCCostoAfectoEntidad, IdCCostoExentoEntidad, IdCCostoTotalEntidad)
            
            'vemos si hay area de negocio o centro de costo para entidad, para marcar como editado el documento, de tal manera que no se genere automáticamente los movimientos y estos datos se pierdan
            If IdAreaNegAfectoEntidad <> 0 Or IdAreaNegExentoEntidad <> 0 Or IdAreaNegTotalEntidad <> 0 Then
               DocEdit = True
            End If
            
            If IdCCostoAfectoEntidad <> 0 Or IdCCostoExentoEntidad <> 0 Or IdCCostoTotalEntidad <> 0 Then
               DocEdit = True
            End If
            
            Descrip = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Diminutivo & " " & vFld(Rs("NumDoc")) & IIf(FCase(NombreEntidad) <> "", " - " & NombreEntidad, "")
            
            Q1 = "UPDATE Documento SET "
            Q1 = Q1 & "  IdCuentaAfecto = " & IdCtaAfecto
            Q1 = Q1 & ", IdCuentaExento = " & IdCtaExento
            Q1 = Q1 & ", IdCuentaTotal = " & IdCtaTotal
            Q1 = Q1 & ", PropIVA = " & IdPropIVA
            Q1 = Q1 & ", Giro = " & IIf(EsDelGiro, 1, 0)
            Q1 = Q1 & ", Descrip = '" & Descrip & "'"
            Q1 = Q1 & ", MovEdited = " & CInt(DocEdit)
            Q1 = Q1 & " WHERE IdDoc = " & IdDoc
            Call ExecSQL(DbMain, Q1)
            
'            If IdEntidad = 98 Then
'               MsgBeep vbExclamation
'            End If
                       
            'Eliminamos el detalle en la db de destino
            Q1 = " WHERE IdDoc = " & IdDoc
            Call DeleteSQL(DbMain, "MovDocumento", Q1)
   
            'Marcamos los movs. de este documento en la tabla de origen, campo EnProceso
            Q1 = "UPDATE " & TmpTblMovDoc
            Q1 = Q1 & " SET EnProceso = 1"
            Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
            Call ExecSQL(DbMain, Q1)
            
            'insertamos los MovDocs marcados en la tabla de destino
            Q1 = "INSERT INTO MovDocumento (IdDoc, Orden, IdCuenta, Debe, Haber, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, Glosa, IdEmpresa, Ano)"
            Q1 = Q1 & " SELECT " & IdDoc & " As IdDoc "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".Orden "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".IdCuenta "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".Debe "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".Haber "
'            Q1 = Q1 & ", " & TmpTblMovDoc & ".Glosa "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".IdTipoValLib "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".EsTotalDoc "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".IdCCosto "
            Q1 = Q1 & ", " & TmpTblMovDoc & ".IdAreaNeg "
            Q1 = Q1 & ", '" & Descrip & "' as Glosa "
            Q1 = Q1 & ", " & gEmpresa.id & " as IdEmpresa "
            Q1 = Q1 & ", " & gEmpresa.Ano & " as Ano "
            Q1 = Q1 & " FROM " & TmpTblMovDoc
            Q1 = Q1 & " WHERE EnProceso <> 0 "
            Call ExecSQL(DbMain, Q1, False)
                        
            'marcamos MovEdited
'            Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc"
'            Q1 = Q1 & " SET Documento.MovEdited = 1"
'            Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib > " & LIBVENTAS_OTROSIMP & " AND Documento.IdDoc = " & IdDoc
'            Call ExecSQL(DbMain, Q1)
            Tbl = " Documento "
            sFrom = " Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc"
            sSet = " Documento.MovEdited = 1"
            sWhere = " WHERE MovDocumento.IdTipoValLib > " & LIBVENTAS_OTROSIMP & " AND Documento.IdDoc = " & IdDoc
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            
            'actualizamos las cuentas si correponde
            If DocEdit Then
               Q1 = "UPDATE MovDocumento "
               Q1 = Q1 & " SET IdCuenta = " & IdCtaAfecto
               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
               Call ExecSQL(DbMain, Q1)
               
               Q1 = "UPDATE MovDocumento "
               Q1 = Q1 & " SET IdCuenta = " & IdCtaExento
               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
               Call ExecSQL(DbMain, Q1)
            
               Q1 = "UPDATE MovDocumento "
               Q1 = Q1 & " SET IdCuenta = " & IdCtaTotal
               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL & " AND IdDoc = " & IdDoc  'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
               Call ExecSQL(DbMain, Q1)
            End If
            
            'ponemos area de negocio y centro de costo por entidad para Afecto si corresponde y si la cuenta tiene ese atributo
            If IdAreaNegAfectoEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdAreaNeg = " & IdAreaNegAfectoEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdAreaNeg = " & IdAreaNegAfectoEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO & " AND IdDoc = " & IdDoc 'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
            
            If IdCCostoAfectoEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdCCosto = " & IdCCostoAfectoEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdCCosto = " & IdCCostoAfectoEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO & " AND IdDoc = " & IdDoc 'LIBVENTAS_AFECTO = LIBCOMPRAS_AFECTO
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
   
            'ponemos area de negocio y centro de costo por entidad para Exento si corresponde y si la cuenta tiene ese atributo
            If IdAreaNegExentoEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdAreaNeg = " & IdAreaNegExentoEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_EXENTO = LIBCOMPRAS_EXENTO
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdAreaNeg = " & IdAreaNegExentoEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO & " AND IdDoc = " & IdDoc   'LIBVENTAS_EXENTO = LIBCOMPRAS_EXENTO
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
            
            If IdCCostoExentoEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdCCosto = " & IdCCostoExentoEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO & " AND IdDoc = " & IdDoc  'LIBVENTAS_EXENTO = LIBCOMPRAS_EXENTO
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdCCosto = " & IdCCostoExentoEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO & " AND IdDoc = " & IdDoc 'LIBVENTAS_EXENTO = LIBCOMPRAS_EXENTO
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
            
            'ponemos area de negocio y centro de costo por entidad para Total si corresponde y si la cuenta tiene ese atributo
            If IdAreaNegTotalEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdAreaNeg = " & IdAreaNegTotalEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL & " AND IdDoc = " & IdDoc  'LIBVENTAS_TOTAL = LIBCOMPRAS_TOTAL
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdAreaNeg = " & IdAreaNegTotalEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL & " AND IdDoc = " & IdDoc 'LIBVENTAS_TOTAL = LIBCOMPRAS_TOTAL
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_AREANEG & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
            
            If IdCCostoExentoEntidad <> 0 Then
'               Q1 = "UPDATE MovDocumento "
'               Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
'               Q1 = Q1 & " SET IdCCosto = " & IdCCostoExentoEntidad
'               Q1 = Q1 & " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL & " AND IdDoc = " & IdDoc  'LIBVENTAS_TOTAL = LIBCOMPRAS_TOTAL
'               Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
'               Call ExecSQL(DbMain, Q1)
               Tbl = " MovDocumento "
               sFrom = " MovDocumento "
               sFrom = sFrom & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
               sSet = " IdCCosto = " & IdCCostoExentoEntidad
               sWhere = " WHERE MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL & " AND IdDoc = " & IdDoc  'LIBVENTAS_TOTAL = LIBCOMPRAS_TOTAL
               sWhere = sWhere & " AND Cuentas.Atrib" & ATRIB_CCOSTO & " <> 0 "
               Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            End If
            
            'Volvemos a limpiar campo EnProceso en la tabla MovDoc de origen
            Q1 = "UPDATE " & TmpTblMovDoc
            Q1 = Q1 & " SET EnProceso = 0 "
            Q1 = Q1 & " WHERE EnProceso <> 0 "
            Call ExecSQL(DbMain, Q1)
   
         End If

      Else
         Call CloseRs(RsDoc)
         
      End If

      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   'actualizamos IdEntidad en los documentos por si hay nuevos y IdEntidad de los DTE no calza con los de Contabilidad
   Tbl = "Documento"
   sFrom = " Documento INNER JOIN Entidades ON Documento.RutEntidad = Entidades.Rut AND Documento.IdEmpresa = Entidades.IdEmpresa"
   sSet = " Documento.IdEntidad = Entidades.IdEntidad "
   sWhere = " WHERE Documento.IdEmpresa = " & gEmpresa.id

   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)


   'eliminamos las tablas temporales
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblDoc)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTblMovDoc)
   
   If NIns > 0 Then
         'eliminamos las cuentas de exento, IVA y OtroImp si el valor del campo es 0
      Q1 = "UPDATE Documento"
      Q1 = Q1 & " SET IdCuentaExento = 0 WHERE Exento = 0 OR Exento IS NULL"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento"
      Q1 = Q1 & " SET IdCuentaIVA = 0 WHERE IVA = 0 OR IVA IS NULL"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento"
      Q1 = Q1 & " SET IdCuentaOtroImp = 0 WHERE OtroImp = 0 OR OtroImp IS NULL"
      Call ExecSQL(DbMain, Q1)
   End If
   
   Call UnLockAction(DbMain, LK_IMPLIBROS, Mes)

'   MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & NIns & " documentos." & vbNewLine & vbNewLine & "- Se actualizaron " & NUpd & " documentos en estado Pendiente.", vbInformation + vbOKOnly
   MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & NIns & " documentos.", vbInformation + vbOKOnly

   ImpFactMes = True
   
  Exit Function
Error_Handler:
    MsgBox "Error #" & ERR.Number & ": '" & ERR.Description & "' from '" & ERR.Source & "'"
    'GoLogTheError sMsg
End Function



Private Sub LoadCuentasDef(ByVal TipoLib As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
   
   lIdCtaAfecto = 0
   lIdCtaExento = 0
   lIdCtaTotal = 0
      
   If TipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion, TipoValor "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = Cuentas.IdEmpresa AND CuentasBasicas.Ano = Cuentas.Ano "
      Q1 = Q1 & " WHERE TipoLib = " & TipoLib
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, CuentasBasicas.Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         Select Case vFld(Rs("TipoValor"))
         
            Case LIBVENTAS_AFECTO, LIBCOMPRAS_AFECTO
            
               If lIdCtaAfecto = 0 Then
                  lIdCtaAfecto = vFld(Rs("IdCuenta"))
               End If
               
            Case LIBVENTAS_EXENTO, LIBCOMPRAS_EXENTO
            
               If lIdCtaExento = 0 Then
                  lIdCtaExento = vFld(Rs("IdCuenta"))
               End If
            
            Case LIBVENTAS_TOTAL, LIBCOMPRAS_TOTAL
            
               If lIdCtaTotal = 0 Then
                  lIdCtaTotal = vFld(Rs("IdCuenta"))
               End If
               
         End Select
                           
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
End Sub

Private Sub GenDbName()
   Dim DbName As String
   Dim ImpName As String, ImpDbPath As String
   Dim StrMes As String, StrAño As String
   Dim Ano As Integer, Mes As Integer, Rut As String
   Dim i As Integer, Idx As Integer
   
   If lInLoad Then
      Exit Sub
   End If
   
   Rut = gEmpresa.Rut
   Ano = Val(Tx_Ano)
   Mes = ItemData(Cb_Mes)
   
   StrAño = Ano
   StrMes = StrAño & Right("0" & Mes, 2)
   
   
   If Op_ImpDesde(IMP_FILE) Then
      ImpName = ReplaceStr(gTipoLib(lTipoLib), "Libro de ", "")
      ImpName = "-" & UCase(Left(ImpName, 3))
   '   If ItemData(Cb_Sucursal) > 0 Then
   '      ImpName = ImpName & "-" & GetCodSucursal(ItemData(Cb_Sucursal))
   '   End If
      ImpName = ImpName & "-" & StrMes
       
      ImpDbPath = gImportPath & "\" & StrAño
'      Rut = 79799800
      If Rut <> "" Then
         DbName = ImpDbPath & "\" & Rut
      Else
         DbName = ImpDbPath & "\" & gEmpresa.Rut
      End If
      
      DbName = DbName & ImpName & ".mdb"
   
      lDbName = DbName
      
      Tx_FName = lDbName
   
   ElseIf gDbFacturacion <> "" Then
   
      Idx = InStrRev(gDbFacturacion, "\")
      If Idx > 0 Then
         DbName = Left(gDbFacturacion, Idx)
         lDbName = DbName & "Empresas\" & gEmpresa.Rut & "-DTE.mdb"
      End If
      
   ElseIf Trim(Tx_DbFactura) <> "" Then
   
      Tx_DbFactura = Trim(Tx_DbFactura)
      Idx = InStrRev(Tx_DbFactura, "\")
      If Idx > 0 Then
         DbName = Left(Tx_DbFactura, Idx)
         lDbName = DbName & "Empresas\" & gEmpresa.Rut & "-DTE.mdb"
      End If
   
   Else
   
      MsgBox1 "Debe seleccionar la ubicación de la base de datos del Sistema de Facturación o Importar desde Archivo exportado en TRFacturación.", vbExclamation
      Exit Sub
   End If
   
End Sub

Private Sub Op_ImpDesde_Click(Index As Integer)
   Call GenDbName
End Sub

Private Sub Op_Libros_Click(Index As Integer)

   lTipoLib = Index

   Call LoadCuentasDef(lTipoLib)
   
   Call GenDbName
End Sub

Private Sub Bt_BrowseDB_Click()
   
   gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = "*.mdb"
   gFrmMain.Cm_ComDlg.Filter = "Datos Facturación|LPFactura*.mdb;TRFactura*.mdb"
   gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Base de Datos de Facturación"
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
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
   
   If LCase(FrmMain.Cm_ComDlg.FileTitle) <> "trfactura.mdb" And LCase(FrmMain.Cm_ComDlg.FileTitle) <> "lpfactura.mdb" Then
      MsgBox1 "Nombre de archivo invalido.", vbExclamation
      Exit Sub
   End If
   
   Tx_DbFactura = FrmMain.Cm_ComDlg.Filename
   gDbFacturacion = Tx_DbFactura
   
   Call GenDbName

End Sub


Private Sub BrowseFile()
   
   gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = "*.mdb"
   gFrmMain.Cm_ComDlg.Filter = "Datos Exportados en TRFacturación|*.mdb"
   gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo Exportado en TRFacturación"
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
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
   
   If Right(FrmMain.Cm_ComDlg.FileTitle, 4) <> ".mdb" Then
      MsgBox1 "Nombre de archivo invalido.", vbExclamation
      Exit Sub
   End If
   
   Tx_FName = FrmMain.Cm_ComDlg.Filename
   gDbFacturacion = Tx_FName
   
   Call GenDbName

End Sub


