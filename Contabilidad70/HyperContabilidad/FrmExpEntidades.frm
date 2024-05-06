VERSION 5.00
Begin VB.Form FrmExpEntidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar Entidades a LPContabilidad"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   975
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   420
      Width           =   2595
      Begin VB.OptionButton Op_Exp 
         Caption         =   "Entidades"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   480
      Picture         =   "FrmExpEntidades.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   690
      TabIndex        =   3
      Top             =   540
      Width           =   690
   End
   Begin VB.CommandButton Bt_Exportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   4860
      TabIndex        =   1
      Top             =   540
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   1020
      Width           =   1575
   End
End
Attribute VB_Name = "FrmExpEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Exportar_Click()
   Dim Mes As Long
   
   Call ExportEnt
   
End Sub

Private Function ExportEnt() As Boolean
   Dim DbName As String
   Dim Db As Database
   Dim ExpName As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim CreateEnable As Boolean
   Dim n As Long
   Dim Msg As String
   Dim i As Integer
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim TblName As String
   Dim UpdOK As Boolean
   Dim Rc As Integer
    
   ExportEnt = False
   
   'Creamos el nombre de la DB de exportación: "Libro-AñoMes.mdb"
   ExpName = "- Entidades"
       
   On Error Resume Next
   MkDir (gExportPath)
   On Error GoTo 0
       
   If ERR Then
      MsgBox1 "Error " & ERR & ": " & Error & " al momento de crear la carpeta de exportación.", vbExclamation
      Exit Function
   End If
   
   DbName = gExportPath & "\" & gEmpresa.Rut & ExpName & ".mdb"

   CreateEnable = LockAction(DbMain, LK_EXPENTIDADES, 0)
   
   If CreateEnable = False Then    'alguien más está exportando entidades
      MsgBox1 "Esta operación ya se está realizando en el equipo '" & IsLockedAction(DbMain, LK_EXPENTIDADES, 0) & "'. No se realizará la exportación.", vbInformation
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
            
   TblName = "Entidades"
   
   Set Tbl = Db.CreateTableDef(TblName)
   
   ERR.Clear
   
   Tbl.Fields.Append Tbl.CreateField("IdEntidad", dbLong)
   Tbl.Fields("IdEntidad").Attributes = dbAutoIncrField
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdEntidad", vbExclamation
      UpdOK = False
   End If
   
   Tbl.Fields.Append Tbl.CreateField("IdEmpresa", dbLong)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdEmpresa", vbExclamation
      UpdOK = False
   End If
   
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Rut", dbText, 12)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Rut", vbExclamation
      UpdOK = False
   End If
        
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Codigo", dbText, 15)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Nombre", dbText, 100)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Nombre", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Direccion", dbText, 100)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Direccion", vbExclamation
      UpdOK = False
   End If
                            
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Region", dbInteger)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Region", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Comuna", dbInteger)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Comuna", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Ciudad", dbText, 20)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Ciudad", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Telefonos", dbText, 30)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Telefonos", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Fax", dbText, 15)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Fax", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("ActEcon", dbLong)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".ActEcon", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("CodActEcon", dbText, 8)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodActEcon", vbExclamation
      UpdOK = False
   End If
                            
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("DomPostal", dbText, 35)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".DomPostal", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("ComPostal", dbInteger)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".ComPostal", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Email", dbText, 100)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Email", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Web", dbText, 50)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Web", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Estado", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Estado", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Obs", dbText, 255)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Obs", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif0", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif0", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif1", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif1", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif2", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif2", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif3", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif3", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif4", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif4", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Clasif5", dbByte)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Clasif5", vbExclamation
      UpdOK = False
   End If
                            
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Giro", dbText, 80)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Giro", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("NotValidRut", dbBoolean)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".NotValidRut", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("EsSupermercado", dbBoolean)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".EsSupermercado", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("EntRelacionada", dbBoolean)
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".EntRelacionada", vbExclamation
      UpdOK = False
   End If
              
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("Ret3Porc", dbBoolean)

   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Ret3Porc", vbExclamation
      UpdOK = False
   End If
              
   Db.TableDefs.Append Tbl
   If ERR = 0 Then
      Db.TableDefs.Refresh
      
      Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (IdEntidad, IdEmpresa)"
      Rc = ExecSQLDao(Db, Q1, False)
      
   ElseIf ERR <> 3010 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
      UpdOK = False
      
   End If
   
   Set Tbl = Nothing
               
   'Insertamos las entidades asociadas a los docs seleccionados
   Q1 = "SELECT * FROM Entidades"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      Q1 = "INSERT INTO Entidades "
      Q1 = Q1 & "(IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut, EsSupermercado, EntRelacionada, Ret3Porc)"
      Q1 = Q1 & " VALUES (" & vFld(Rs("IdEmpresa")) & ",'" & vFld(Rs("Rut")) & "','" & vFld(Rs("Codigo")) & "','" & ParaSQL(vFld(Rs("Nombre"))) & "','" & ParaSQL(vFld(Rs("Direccion"))) & "'," & vFld(Rs("Region"))
      Q1 = Q1 & "," & vFld(Rs("Comuna")) & ",'" & vFld(Rs("Ciudad")) & "','" & vFld(Rs("Telefonos")) & "','" & vFld(Rs("Fax")) & "'," & vFld(Rs("ActEcon")) & ",'" & vFld(Rs("CodActEcon")) & "'"
      Q1 = Q1 & ",'" & vFld(Rs("DomPostal")) & "'," & vFld(Rs("ComPostal")) & ",'" & vFld(Rs("Email")) & "','" & vFld(Rs("Web")) & "'," & vFld(Rs("Estado")) & ",'" & vFld(Rs("Obs")) & "'"
      Q1 = Q1 & "," & vFld(Rs("Clasif0")) & "," & vFld(Rs("Clasif1")) & "," & vFld(Rs("Clasif2")) & "," & vFld(Rs("Clasif3")) & "," & vFld(Rs("Clasif4")) & "," & vFld(Rs("Clasif5"))
      Q1 = Q1 & ",'" & vFld(Rs("Giro")) & "'," & Abs(vFld(Rs("NotValidRut"))) & "," & Abs(vFld(Rs("EsSupermercado"))) & "," & Abs(vFld(Rs("EntRelacionada"))) & "," & Abs(vFld(Rs("Ret3Porc"))) & ")"
      
      Call ExecSQLDao(Db, Q1)
   
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   'vemos cuántos docs se exportaron
   Q1 = "SELECT Count(*) FROM Entidades"
   Set RsDao = OpenRsDao(Db, Q1)
   n = RsDao(0)
   Call CloseRs(RsDao)
   
   Select Case n
      Case 0
         Msg = "No se encontraron entidades para exportar."
      Case 1
         Msg = "Se exportó una Entidad." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
      Case Else
         Msg = "Se exportaron " & n & " Entidades." & vbNewLine & vbNewLine
         Msg = Msg & "Archivo generado:" & vbNewLine & vbNewLine
         Msg = Msg & "      " & DbName
   End Select
   
      
   Call CloseDb(Db)
   
   Call UnLockAction(DbMain, LK_EXPENTIDADES, 0)
   
   MsgBox1 Msg, vbInformation + vbOKOnly
   
   ExportEnt = True

End Function

