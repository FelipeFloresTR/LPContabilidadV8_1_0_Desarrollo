VERSION 5.00
Begin VB.Form FrmImpExpLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar/Exportar Libros Auxiliares desde Sucursal"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Ch_Plan 
      Caption         =   "Exigir que el plan de cuentas del origen sea igual al plan de cuentas del destino."
      Height          =   435
      Left            =   1440
      TabIndex        =   15
      Top             =   3720
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CommandButton Bt_Export 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   1635
      Left            =   1380
      TabIndex        =   9
      Top             =   1860
      Width           =   4395
      Begin VB.ComboBox Cb_Sucursal 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   3075
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   11
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Index           =   0
      Left            =   1380
      TabIndex        =   8
      Top             =   420
      Width           =   2595
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   2235
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Retenciones"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Width           =   2355
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   420
      Picture         =   "FrmImpExpLib.frx":0000
      ScaleHeight     =   660
      ScaleWidth      =   660
      TabIndex        =   7
      Top             =   540
      Width           =   660
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4500
      TabIndex        =   6
      Top             =   900
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   4500
      TabIndex        =   5
      Top             =   540
      Width           =   1275
   End
End
Attribute VB_Name = "FrmImpExpLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lMsgAdv As Boolean
Dim lRc As Integer
Dim lOper As Integer
Dim lQryCtas As String

Private Sub Bt_Cancelar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim i As Integer
   
   If Ch_Plan.Value = 0 Then
   
      If MsgBox1("Si no se valida la igualdad del plan de cuentas, es posible que algunas cuentas del libro queden en blanco." & vbNewLine & vbNewLine & "¿Desea continuar con la importación?", vbQuestion + vbYesNo) = vbNo Then
         Exit Sub
      End If
      
   End If
   
  
   For i = LIB_COMPRAS To LIB_RETEN
   
      If Op_Libros(i).Value = True Then
   
         If lMsgAdv = False Then    'este mensaje se muestra sólo una vez
         
            If MsgBox1("Para realizar la importación del " & gTipoLib(i) & ", nadie más debe estar trabajando en este libro para esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         
         End If
   
         lMsgAdv = True
         
         Me.MousePointer = vbHourglass
         Call ImportLibroMes(i, gEmpresa.Rut, Val(Tx_Ano), ItemData(Cb_Mes))
         Me.MousePointer = vbDefault
         
         Exit For
         
      End If
   Next i
   

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim Q1 As String
   Dim i As Integer
   
   lMsgAdv = False
   
   If lOper = O_IMPORT Then
      Me.Caption = "Importar Libro Auxiliar desde Sucursal"
      Bt_Export.visible = False
   Else
      Me.Caption = "Exportar Libro Auxiliar desde Sucursal"
      Bt_Import.visible = False
      Ch_Plan.visible = False
      Ch_Plan = 1
   End If
   
   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   
   Tx_Ano = gEmpresa.Ano
   
   Call AddItem(Cb_Sucursal, " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id               '   & " AND Vigente <> 0 "
   Q1 = Q1 & " ORDER BY Descripcion"
   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)

   'generamos Query para checksum de Plan de Cuentas
   lQryCtas = "SELECT Codigo, Descripcion, Nivel, Clasificacion "
'   For i = 1 To 10
'      lQryCtas = lQryCtas & ", Atrib" & i
'   Next i
   
   lQryCtas = lQryCtas & " FROM Cuentas"
   lQryCtas = lQryCtas & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

End Sub
Public Function FImport() As Integer

   lOper = O_IMPORT
    Me.Show vbModal

End Function
Public Function FExport() As Integer

   lOper = O_EXPORT
    Me.Show vbModal

End Function
Private Sub Bt_Export_Click()
   Dim i As Integer
   
   
   For i = LIB_COMPRAS To LIB_RETEN
   
      If Op_Libros(i).Value = True Then
   
         If lMsgAdv = False Then    'este mensaje se muestra sólo una vez
         
            If MsgBox1("Para realizar la exportación del " & gTipoLib(i) & " de " & Cb_Mes & ", nadie más debe estar trabajando en este libro para esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         
         End If
   
         lMsgAdv = True
         
         Me.MousePointer = vbHourglass
         Call ExportLibroMes(i, gEmpresa.Rut, Val(Tx_Ano), ItemData(Cb_Mes))
         Me.MousePointer = vbDefault
         
         Exit For
         
      End If
      
   Next i
   

End Sub

Private Function ExportLibroMes(ByVal TipoLib As Integer, ByVal Rut As String, ByVal Ano As Integer, ByVal Mes As Integer) As Boolean
   Dim DbName As String
   Dim Db As Database
   Dim LibExpName As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim StrMes As String
   Dim CreateEnable As Boolean
   Dim n As Long
   Dim ExpDbPath As String
   Dim Msg As String
   Dim StrAño As String
   Dim i As Integer
   Dim ChkSumCuentas As Long
   Dim Tbl As TableDef
   Dim Fld As Field
    
   If TipoLib = 0 Or Mes = 0 Then
      Exit Function
   End If
   
   ExportLibroMes = False
   
   StrMes = Right("0" & Mes, 2)
   StrAño = Right(gEmpresa.Ano, 2)
   
   'Creamos el nombre de la DB de exportación: "Libro-AñoMes.mdb"
   LibExpName = ReplaceStr(gTipoLib(TipoLib), "Libro de ", "")
   LibExpName = "-" & UCase(Left(LibExpName, 3))
   If ItemData(Cb_Sucursal) > 0 Then
      LibExpName = LibExpName & "-" & GetCodSucursal(ItemData(Cb_Sucursal))
   End If
   LibExpName = LibExpName & "-" & StrAño & StrMes
    
   ExpDbPath = gExportPath & "\Libros\"
   
   On Error Resume Next
   MkDir ExpDbPath
   
   On Error GoTo 0
       
   If Ano > 0 Then
      If Rut <> "" Then
         DbName = ExpDbPath & "\" & Rut & LibExpName & ".mdb"
      Else
         DbName = ExpDbPath & "\" & gEmpresa.Rut & LibExpName & ".mdb"
      End If
      
   ElseIf Rut <> "" Then
      DbName = ExpDbPath & "\" & Rut & LibExpName & ".mdb"
   Else
      DbName = ExpDbPath & "\" & gEmpresa.Rut & LibExpName & ".mdb"
   End If

   CreateEnable = LockAction(DbMain, LK_EXPLIBROS, Mes)
   
   If CreateEnable = False Then    'alguien más está exportando este mes
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_EXPLIBROS, Mes) & "'. No se realizará la exportación.", vbInformation
      Exit Function
   End If
   
   On Error Resume Next
   
   Kill (DbName)
   ERR.Clear
   
   'creamos la DB
   Set Db = CreateDatabase(DbName, dbLangGeneral)
      
   If (ERR Or Db Is Nothing) And ERR <> 3204 Then
      MsgBox "Error " & ERR & ", " & Error & NL & DbName, vbExclamation
      Db.Close
      Set Db = Nothing
      Exit Function
   End If
   
   On Error GoTo 0

#If DATACON = 1 Then

   'linkeamos a la DB las tablas que necesitamos
   Call LinkMdbTable(Db, DbMain.Name, "Documento", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "MovDocumento", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "Entidades", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "Cuentas", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "Sucursales", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "AreaNegocio", , , , gEmpresa.ConnStr)
   Call LinkMdbTable(Db, DbMain.Name, "CentroCosto", , , , gEmpresa.ConnStr)
   
#End If

   'generamos checksum de Plan de Cuentas, para verificarlo al importar
   ChkSumCuentas = ChkSumQry(DbMain, lQryCtas)
   
   Set Tbl = New TableDef
   Tbl.Name = "ParamExp"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("ChkSumCuentas", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   Else
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ParamExp.ChkSumCuentas", vbExclamation
   End If
      
   Db.TableDefs.Append Tbl
   If ERR = 0 Then
      DbMain.TableDefs.Refresh
      
      Q1 = "INSERT INTO ParamExp (ChkSumCuentas) VALUES(" & ChkSumCuentas & ")"
      Call ExecSQL(Db, Q1)
      
   Else
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla ParamExp", vbExclamation
      
   End If
   
   Set Tbl = Nothing
   
   'generamos los registros del libro-mes
   Q1 = "SELECT Documento.*, Entidades.RUT, Entidades.Codigo, Entidades.Nombre, Entidades.NotValidRut,"
   Q1 = Q1 & " Cuentas1.Codigo as CodCtaEx, Cuentas2.Codigo as CodCtaAf, Cuentas3.Codigo as CodCtaIVA, "
   Q1 = Q1 & " Cuentas4.Codigo As CodCtaOtroImp, Cuentas5.Codigo as CodCtaTotal,Cuentas6.Codigo as CodCta3porc, "
   Q1 = Q1 & " Sucursales.Codigo As CodSuc  "
   Q1 = Q1 & " INTO Documento" & StrMes
   'se añade una sexta cuenta para idCuentaRet3porc gcb21102021
   Q1 = Q1 & " FROM (((((((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad) "
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaAfecto = Cuentas2.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas3 ON Documento.IdCuentaIVA = Cuentas3.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas4 ON Documento.IdCuentaOtroImp = Cuentas4.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas5 ON Documento.IdCuentaTotal = Cuentas5.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas6 ON Documento.idCuentaRet3porc = Cuentas6.IdCuenta )"
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal "
   
   Q1 = Q1 & " WHERE " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   Q1 = Q1 & " AND Documento.TipoLib = " & TipoLib
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   If ItemData(Cb_Sucursal) > 0 Then
      Q1 = Q1 & " AND Documento.IdSucursal = " & ItemData(Cb_Sucursal)
   End If
   
   Call ExecSQL(Db, Q1)
      
   'Insertamos los detalles de los documentos seleccionados
   Q1 = "SELECT MovDocumento.IdDoc, MovDocumento.Orden, 0 as IdCuenta, MovDocumento.Debe"
   Q1 = Q1 & ", MovDocumento.Haber, MovDocumento.Glosa, MovDocumento.IdTipoValLib"
   Q1 = Q1 & ", MovDocumento.EsTotalDoc, 0 as IdAreaNeg, 0 as IdCCosto "
   Q1 = Q1 & ", Cuentas.Codigo As CodCta, AreaNegocio.Codigo as CodAreaNeg "
   Q1 = Q1 & ", CentroCosto.Codigo as CodCCosto, 0 as EnProceso "
   Q1 = Q1 & ", " & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano "
   Q1 = Q1 & " INTO MovDocumento" & StrMes
   Q1 = Q1 & " FROM ((( MovDocumento INNER JOIN Documento" & StrMes & " ON MovDocumento.IdDoc = Documento" & StrMes & ".IdDoc )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta ) "
   Q1 = Q1 & " LEFT JOIN AreaNegocio ON MovDocumento.IdAreaNeg = AreaNegocio.IdAreaNegocio) "
   Q1 = Q1 & " LEFT JOIN CentroCosto ON MovDocumento.IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & " WHERE MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
   Call ExecSQL(Db, Q1)
   
   'Insertamos las entidades asociadas a los docs seleccionados
   Q1 = "SELECT DISTINCT Entidades.Rut, Entidades.Codigo, Entidades.Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal"
   Q1 = Q1 & ", ComPostal, Email, Web, Entidades.Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Entidades.Giro, Entidades.NotValidRut"
   Q1 = Q1 & ", EsSupermercado, Entidades.IdEmpresa, Entidades.EntRelacionada, CodCtaAfecto, CodCtaExento, Entidades.CodCtaTotal, Entidades.PropIVA, CodCCostoAfecto"
   Q1 = Q1 & ", CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, Entidades.CodCtaTotalVta "
   Q1 = Q1 & ", CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta, EsDelGiro   "
   Q1 = Q1 & " INTO Entidades" & StrMes
   Q1 = Q1 & " FROM Entidades INNER JOIN Documento" & StrMes & " ON Entidades.IdEntidad = Documento" & StrMes & ".IdEntidad "
   Q1 = Q1 & " WHERE Entidades.IdEmpresa = " & gEmpresa.id
   Call ExecSQL(Db, Q1)
   
   'limpiamos algunos campos por si acaso (esto debe ir después de Insert de entidades porque necesitamos el IdEntidad)
   Q1 = "UPDATE Documento" & StrMes & " SET "
   Q1 = Q1 & "  IdCompCent = 0"
   Q1 = Q1 & ", IdCompPago = 0"
   Q1 = Q1 & ", SaldoDoc = 0"
   Q1 = Q1 & ", FExported = 0"    'esta fecha se pone cuando se lleva un doc de un año a otro
   Q1 = Q1 & ", OldIdDoc = 0"     'este campo se pone cuando se lleva un doc de un año a otro
   Q1 = Q1 & ", FImporF29 = 0"
   Q1 = Q1 & ", NumDocRef = '0'"    'este campo se usa al importar desde Form29
   Q1 = Q1 & ", IdCtaBanco = 0"
   Q1 = Q1 & ", IdUsuario = 0"
   Q1 = Q1 & ", IdEntidad = 0"
   Q1 = Q1 & ", IdCuentaExento = 0"
   Q1 = Q1 & ", IdCuentaAfecto = 0"
   Q1 = Q1 & ", IdCuentaIVA = 0"
   Q1 = Q1 & ", IdCuentaOtroImp = 0"
   Q1 = Q1 & ", IdSucursal = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(Db, Q1)
   
   'vemos cuántos docs se exportaron
   Q1 = "SELECT Count(*) FROM Documento" & StrMes
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(Db, Q1)
   n = Rs(0)
   Call CloseRs(Rs)
   
   Select Case n
      Case 0
         Msg = "No se encontraron documentos para exportar."
      Case 1
         Msg = "Se exportó un documento." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
      Case Else
         Msg = "Se exportaron " & n & " documentos." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
   End Select
   
#If DATACON = 1 Then
   
   Call UnLinkTable(Db, "Documento")
   Call UnLinkTable(Db, "MovDocumento")
   Call UnLinkTable(Db, "Entidades")
   Call UnLinkTable(Db, "Cuentas")
   Call UnLinkTable(Db, "Sucursales")
   Call UnLinkTable(Db, "AreaNegocio")
   Call UnLinkTable(Db, "CentroCosto")
#End If
   
   Call CloseDb(Db)
   
   Call UnLockAction(DbMain, LK_EXPLIBROS, Mes)
   
   MsgBox1 Msg, vbInformation + vbOKOnly
   
   ExportLibroMes = True
   
End Function
Private Function ImportLibroMes(ByVal TipoLib As Integer, ByVal Rut As String, ByVal Ano As Integer, ByVal Mes As Integer) As Boolean
   Dim DbName As String
   Dim Db As Database
   Dim LibExpName As String
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
   Dim FldName As String
   Dim NUpd As Long
   Dim NIns As Long
   Dim ImpDbPath As String
   Dim StrAño As String
   Dim ChkSumCtasLoc As Long
   Dim ChkSumCtasImp As Long
   Dim Rc As Integer
   Dim FldArray(4) As AdvTbAddNew_t
   Dim FldUpd As String
   
   If TipoLib = 0 Or Mes = 0 Then
      Exit Function
   End If
   
   ImportLibroMes = False
   
   StrMes = Right("0" & Mes, 2)
   StrAño = Right(gEmpresa.Ano, 2)
   
   LibExpName = ReplaceStr(gTipoLib(TipoLib), "Libro de ", "")
   LibExpName = "-" & UCase(Left(LibExpName, 3))
   If ItemData(Cb_Sucursal) > 0 Then
      LibExpName = LibExpName & "-" & GetCodSucursal(ItemData(Cb_Sucursal))
   End If
   LibExpName = LibExpName & "-" & StrAño & StrMes
    
   ImpDbPath = gImportPath
       
   If Ano > 0 Then
      If Rut <> "" Then
         DbName = ImpDbPath & "\" & Rut & LibExpName & ".mdb"
      Else
         DbName = ImpDbPath & "\" & gEmpresa.Rut & LibExpName & ".mdb"
      End If
      
   ElseIf Rut <> "" Then
      DbName = ImpDbPath & "\" & Rut & LibExpName & ".mdb"
   Else
      DbName = ImpDbPath & "\" & gEmpresa.Rut & LibExpName & ".mdb"
   End If
   
   If Not ExistFile(DbName) Then
      MsgBox1 "No se encontró el archivo:" & vbNewLine & vbNewLine & "        " & DbName & vbNewLine & vbNewLine & "Verifique que el archivo se encuentre en la carpeta especificada y vuelva a intentarlo.", vbExclamation + vbOKOnly
      Exit Function
   End If

   ImpEnable = LockAction(DbMain, LK_IMPLIBROS, Mes)
   
   If ImpEnable = False Then    'alguien más está importando este mes
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_EXPLIBROS, Mes) & "'. No se realizará la importación.", vbInformation
      Exit Function
   End If
   
   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & DbName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
   
#If DATACON = 1 Then
   
   'la base de datos de exportación no tiene password
   Call LinkMdbTable(DbMain, DbName, "Documento" & StrMes)
   Call LinkMdbTable(DbMain, DbName, "MovDocumento" & StrMes)
   Call LinkMdbTable(DbMain, DbName, "Entidades" & StrMes)
   Call LinkMdbTable(DbMain, DbName, "ParamExp")
   
#End If

   'Primero Revisamos si hay diferencias en los planes de cuenta con el CheckSum que viene en la base de datos de importación
   Q1 = "SELECT ChkSumCuentas FROM ParamExp"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      ChkSumCtasImp = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   ChkSumCtasLoc = ChkSumQry(DbMain, lQryCtas)
   
   If ChkSumCtasLoc <> ChkSumCtasImp Then
      If Ch_Plan <> 0 Then
         MsgBox1 "El plan de cuentas local es distinto del plan de cuentas definido en el lugar donde se realizó la exportación del libro." & vbNewLine & vbNewLine & "Las diferencias pueden estar en los códigos de las cuentas, la descripción, la clasificación o los atributos." & vbNewLine & vbNewLine & "Realice las modificaciones correspondientes y vuelva a intentarlo.", vbExclamation + vbOKOnly
         Exit Function
      
      ElseIf MsgBox1("El plan de cuentas local es distinto del plan de cuentas definido en el lugar donde se realizó la exportación del libro." & vbNewLine & vbNewLine & "Las diferencias pueden estar en los códigos de las cuentas, la descripción, la clasificación o los atributos." & vbNewLine & vbNewLine & "Algunas cuentas del libro podrían quedar en blanco ¿Desea continuar con el proceso de importación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      
      End If
   End If
   
   'primero traemos las entidades nuevas (dado los índices definidos en la tabla Entidades, sólo se insertarán los Ruts y Códigos nuevos (son únicos))
'   Q1 = "INSERT INTO Entidades "
'   Q1 = Q1 & "  ( Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal"
'   Q1 = Q1 & ", ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut"
'   Q1 = Q1 & ", EsSupermercado, IdEmpresa, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto"
'   Q1 = Q1 & ", CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta "
'   Q1 = Q1 & ", EsDelGiro, CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta ) "
'   Q1 = Q1 & "  SELECT Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal"
'   Q1 = Q1 & ", ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut"
'   Q1 = Q1 & ", EsSupermercado, IdEmpresa, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto"
'   Q1 = Q1 & ", CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, Entidades.CodCtaTotalVta "
'   Q1 = Q1 & ", Entidades.Giro, CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta   "
'
'   Q1 = Q1 & " FROM Entidades" & StrMes
'   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
'   Call ExecSQL(DbMain, Q1)
      
   Q1 = "INSERT INTO Entidades "
   Q1 = Q1 & "  ( Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal"
   Q1 = Q1 & ", ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut"
   Q1 = Q1 & ", EsSupermercado, IdEmpresa, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto"
   Q1 = Q1 & ", CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta "
   Q1 = Q1 & ", EsDelGiro, CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta ) "
   
   Q1 = Q1 & "  SELECT Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal"
   Q1 = Q1 & ", ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut"
   Q1 = Q1 & ", EsSupermercado, IdEmpresa, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto"
   Q1 = Q1 & ", CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta "
   Q1 = Q1 & ", EsDelGiro,CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta   "

   Q1 = Q1 & " FROM Entidades" & StrMes
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)
      
      
      
   'actualizamos el Id de la entidad en los registros de origen
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Entidades ON Documento" & StrMes & ".Rut =  Entidades.Rut"
   Q1 = Q1 & " AND Entidades.IdEmpresa = Documento" & StrMes & ".IdEmpresa "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos el Id de la Sucursal en los registros de origen
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Sucursales ON Documento" & StrMes & ".CodSuc =  Sucursales.Codigo"
   Q1 = Q1 & " AND Sucursales.IdEmpresa = Documento" & StrMes & ".IdEmpresa "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdSucursal = Sucursales.IdSucursal "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'dejamos en cero los IdCuenta de Documentos
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaExento = 0, "
   Q1 = Q1 & "     Documento" & StrMes & ".IdCuentaAfecto = 0, "
   Q1 = Q1 & "     Documento" & StrMes & ".IdCuentaIVA = 0, "
   Q1 = Q1 & "     Documento" & StrMes & ".IdCuentaOtroImp = 0, "
   Q1 = Q1 & "     Documento" & StrMes & ".IdCuentaTotal = 0 "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Actualizamos los Ids de las  cuentas en los registros de origen
   'Exento
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento" & StrMes & ".CodCtaEx = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = Documento" & StrMes & ".IdEmpresa AND Cuentas.Ano = Documento" & StrMes & ".Ano "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaExento = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Afecto
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento" & StrMes & ".CodCtaAf = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = Documento" & StrMes & ".IdEmpresa AND Cuentas.Ano = Documento" & StrMes & ".Ano "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaAfecto = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'IVA
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento" & StrMes & ".CodCtaIVA = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = Documento" & StrMes & ".IdEmpresa AND Cuentas.Ano = Documento" & StrMes & ".Ano "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaIVA = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'OtroImp
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento" & StrMes & ".CodCtaOtroImp = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = Documento" & StrMes & ".IdEmpresa AND Cuentas.Ano = Documento" & StrMes & ".Ano "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaOtroImp = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Total
   Q1 = "UPDATE Documento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON Documento" & StrMes & ".CodCtaTotal = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = Documento" & StrMes & ".IdEmpresa AND Cuentas.Ano = Documento" & StrMes & ".Ano "
   Q1 = Q1 & " SET Documento" & StrMes & ".IdCuentaTotal = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE Documento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND Documento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'limpiamos el campo EnProceso en MovDoc de origen, por si hubo un corte abrupto del
   'proceso de importación y ahora se está volviendo a ejecutar la importación
   Q1 = "UPDATE MovDocumento" & StrMes
   Q1 = Q1 & " SET EnProceso = 0 "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
    
   'Actualizamos el Id de la Cuenta en los MovDoc de origen (primero lo dejamos en cero)
   Q1 = "UPDATE MovDocumento" & StrMes
   Q1 = Q1 & " SET MovDocumento" & StrMes & ".IdCuenta = 0 "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE MovDocumento" & StrMes
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento" & StrMes & ".CodCta = Cuentas.Codigo "
   Q1 = Q1 & " AND Cuentas.IdEmpresa = MovDocumento" & StrMes & ".IdEmpresa AND Cuentas.Ano = MovDocumento" & StrMes & ".Ano "
   Q1 = Q1 & " SET MovDocumento" & StrMes & ".IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & " WHERE MovDocumento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND MovDocumento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
 
   'Actualizamos el Id del Area de Negocio en los MovDoc de origen
   Q1 = "UPDATE MovDocumento" & StrMes
   Q1 = Q1 & " INNER JOIN AreaNegocio ON MovDocumento" & StrMes & ".CodAreaNeg = AreaNegocio.Codigo "
   Q1 = Q1 & " AND AreaNegocio.IdEmpresa = MovDocumento" & StrMes & ".IdEmpresa "
   Q1 = Q1 & " SET MovDocumento" & StrMes & ".IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & " WHERE MovDocumento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND MovDocumento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
 
   'Actualizamos el Id del Centro de Costo en los MovDoc de origen
   Q1 = "UPDATE MovDocumento" & StrMes
   Q1 = Q1 & " INNER JOIN CentroCosto ON MovDocumento" & StrMes & ".CodCCosto = CentroCosto.Codigo "
   Q1 = Q1 & " AND CentroCosto.IdEmpresa = MovDocumento" & StrMes & ".IdEmpresa "
   Q1 = Q1 & " SET MovDocumento" & StrMes & ".IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & " WHERE MovDocumento" & StrMes & ".IdEmpresa = " & gEmpresa.id & " AND MovDocumento" & StrMes & ".Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
 
   Q1 = "SELECT * FROM Documento" & StrMes
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   NUpd = 0
   NIns = 0
   
'   If Not Rs.EOF Then
'      'armamos el UPDATE campo a campo
'      FldUpd = ""
'
'      For i = 0 To Rs.Fields.Count - 1
'         FldName = UCase(Rs.Fields(i).Name)
'         If FldName = "RUT" Then   'inicio de campos adicionales de exportación
'            Exit For
'         End If
'         If FldName <> "IDDOC" And FldName <> "ESTADO" And FldName <> "FECHACREACION" And FldName <> "IDUSUARIO" And FldName <> "FIMPORTSUC" Then
'            If FldIsString(Rs.Fields(i)) Then
'               FldUpd = FldUpd & "," & Rs.Fields(i).Name & "= '" & ParaSQL(vFld(Rs(Rs.Fields(i).Name))) & "'"
'            Else
'               FldUpd = FldUpd & "," & Rs.Fields(i).Name & "= " & vFld(Rs(Rs.Fields(i).Name))
'            End If
'         End If
'      Next i
'
'      FldUpd = FldUpd & ", FImportSuc=" & CLng(Int(Now))   'guardamos la fecha de la última actualización
'
'   End If
   
'   FldUpd = Mid(FldUpd, 2)  'sacamos la primera coma
   
   'ahora importamos uno a uno
   Do While Not Rs.EOF
   
         FldUpd = ""
      
      For i = 0 To Rs.Fields.Count - 1
         FldName = UCase(Rs.Fields(i).Name)
         If FldName = "RUT" Then   'inicio de campos adicionales de exportación
            Exit For
         End If
         If FldName <> "IDDOC" And FldName <> "ESTADO" And FldName <> "FECHACREACION" And FldName <> "IDUSUARIO" And FldName <> "FIMPORTSUC" Then
            If FldIsString(Rs.Fields(i)) Then
               FldUpd = FldUpd & "," & Rs.Fields(i).Name & "= '" & ParaSQL(vFld(Rs(Rs.Fields(i).Name))) & "'"
            Else
               FldUpd = FldUpd & "," & Rs.Fields(i).Name & "= " & vFld(Rs(Rs.Fields(i).Name))
            End If
         End If
      Next i
      
      
      FldUpd = FldUpd & ", FImportSuc=" & CLng(Int(Now))
      FldUpd = Mid(FldUpd, 2)
   
      Q1 = "SELECT * "
      Q1 = Q1 & " FROM Documento "
      Q1 = Q1 & " WHERE TipoLib = " & vFld(Rs("TipoLib"))
      Q1 = Q1 & " AND TipoDoc = " & vFld(Rs("TipoDoc"))
      Q1 = Q1 & " AND NumDoc = '" & vFld(Rs("NumDoc")) & "'"
      
      If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
         Q1 = Q1 & " AND NumDocHasta = '" & vFld(Rs("NumDocHasta")) & "'"
      Else
         Q1 = Q1 & " AND (NumDocHasta = '0' OR NumDocHasta IS NULL) "
      End If
      
      If vFld(Rs("IdEntidad")) <> 0 Then
         Q1 = Q1 & " AND Documento.IdEntidad = " & vFld(Rs("IdEntidad"))
      Else
         Q1 = Q1 & " AND (Documento.IdEntidad IS NULL OR Documento.IdEntidad = 0)"
      End If
   
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      'obtenemos IdDoc
      Set RsDoc = OpenRs(DbMain, Q1)
      
      IdDoc = 0
      
      If RsDoc.EOF Then   'no está, lo agregamos
      
         'insertamos el Doc
'         Set RsNewDoc = DbMain.OpenRecordset("Documento")
'         RsNewDoc.AddNew
'
'         IdDoc = vFld(RsNewDoc("IdDoc"))
'
'         RsNewDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsNewDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         If vFld(Rs("Estado")) = ED_ANULADO Then
'            RsNewDoc.Fields("Estado") = ED_ANULADO
'         Else                           'pagados o centralizados en la sucursal se dejan igual pendientes en la central
'            RsNewDoc.Fields("Estado") = ED_PENDIENTE
'         End If
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsNewDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsNewDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         RsNewDoc.Close
'
'         Set RsNewDoc = Nothing
                  
         FldArray(0).FldName = "IdUsuario"
         FldArray(0).FldValue = gUsuario.IdUsuario
         FldArray(0).FldIsNum = True
         
         FldArray(1).FldName = "FechaCreacion"
         FldArray(1).FldValue = CLng(Int(Now))
         FldArray(1).FldIsNum = True
         
         FldArray(2).FldName = "Estado"
         If vFld(Rs("Estado")) = ED_ANULADO Then
            FldArray(2).FldValue = ED_ANULADO
         Else                                      'pagados o centralizados en la sucursal se dejan igual pendientes en la central
            FldArray(2).FldValue = ED_PENDIENTE
         End If
         FldArray(2).FldIsNum = True
         
         FldArray(3).FldName = "IdEmpresa"
         FldArray(3).FldValue = gEmpresa.id
         FldArray(3).FldIsNum = True
                     
         FldArray(4).FldName = "Ano"
         FldArray(4).FldValue = gEmpresa.Ano
         FldArray(4).FldIsNum = True
                     
         IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
                  
         NIns = NIns + 1
         
      ElseIf vFld(RsDoc("Estado")) = ED_PENDIENTE Then   'ya existe y está pendiente, lo actualizamos
      
         IdDoc = vFld(RsDoc("IdDoc"))
         NUpd = NUpd + 1
         
      End If
         
      Call CloseRs(RsDoc)
      
      'actualizamos el doc
      
      If IdDoc <> 0 Then
      
         'armamos el UPDATE campo a campo
'         Q1 = ""
'
'         For i = 0 To Rs.Fields.Count - 1
'            FldName = UCase(Rs.Fields(i).Name)
'            If FldName = "RUT" Then   'inicio de campos adicionales de exportación
'               Exit For
'            End If
'            If FldName <> "IDDOC" And FldName <> "ESTADO" And FldName <> "FECHACREACION" And FldName <> "IDUSUARIO" And FldName <> "FIMPORTSUC" Then
'               If FldIsString(Rs.Fields(i)) Then
'                  Q1 = Q1 & "," & Rs.Fields(i).Name & "= '" & ParaSQL(vFld(Rs(Rs.Fields(i).Name))) & "'"
'               Else
'                  Q1 = Q1 & "," & Rs.Fields(i).Name & "= " & vFld(Rs(Rs.Fields(i).Name))
'               End If
'            End If
'         Next i
'
'         Q1 = Q1 & ", FImportSuc=" & CLng(Int(Now))   'guardamos la fecha de la última actualización
         
         Q1 = "UPDATE Documento SET " & FldUpd
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         'Eliminamos el detalle en la db de destino
'         Q1 = "DELETE * FROM MovDocumento WHERE IdDoc = " & IdDoc
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "MovDocumento", Q1)
         
         'Marcamos los movs. de este documento en la tabla de origen, campo EnProceso
         Q1 = "UPDATE MovDocumento" & StrMes
         Q1 = Q1 & " SET EnProceso = 1"
         Q1 = Q1 & " WHERE IdDoc = " & vFld(Rs("IdDoc"))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocs marcados en la tabla de destino
         Q1 = "INSERT INTO MovDocumento "
         Q1 = Q1 & "( IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, IdEmpresa, Ano )"
         Q1 = Q1 & " SELECT " & IdDoc & " As IdDoc "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".Orden "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".IdCuenta "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".Debe "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".Haber "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".Glosa "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".IdTipoValLib "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".EsTotalDoc "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".IdCCosto "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".IdAreaNeg "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".IdEmpresa "
         Q1 = Q1 & " , MovDocumento" & StrMes & ".Ano "
         Q1 = Q1 & " FROM MovDocumento" & StrMes
         Q1 = Q1 & " WHERE EnProceso <> 0 "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
                  
         'Volvemos a limpiar campo EnProceso en la tabla MovDoc de origen
         Q1 = "UPDATE MovDocumento" & StrMes
         Q1 = Q1 & " SET EnProceso = 0 "
         Q1 = Q1 & " WHERE EnProceso <> 0 "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      'Tracking 3227543
     Call SeguimientoDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "FrmImpExpLib.ImportLibroMes", "", 1, "", gUsuario.IdUsuario, 2, 1)
     Call SeguimientoMovDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "FrmImpExpLib.ImportLibroMes", "", 1, "", 2, 1)
     ' fin 3227543
      
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
#If DATACON = 1 Then
   Call UnLinkTable(DbMain, "Documento" & StrMes)
   Call UnLinkTable(DbMain, "MovDocumento" & StrMes)
   Call UnLinkTable(DbMain, "Entidades" & StrMes)
   Call UnLinkTable(DbMain, "ParamExp")
#End If

   Call UnLockAction(DbMain, LK_IMPLIBROS, Mes)
   
   MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & NIns & " documentos." & vbNewLine & vbNewLine & "- Se actualizaron " & NUpd & " documentos en estado Pendiente.", vbInformation + vbOKOnly

   ImportLibroMes = True
   
End Function


