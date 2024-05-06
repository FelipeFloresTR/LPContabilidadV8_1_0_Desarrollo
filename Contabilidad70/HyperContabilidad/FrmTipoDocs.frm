VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmTipoDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "FrmTipoDocs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   420
      Picture         =   "FrmTipoDocs.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   8
      Top             =   600
      Width           =   585
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5700
      TabIndex        =   3
      Top             =   540
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipos de Documentos"
      Height          =   3795
      Left            =   1260
      TabIndex        =   7
      Top             =   1380
      Width           =   5715
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   855
         Left            =   4440
         Picture         =   "FrmTipoDocs.frx":0551
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar Entidad"
         Top             =   420
         Width           =   1095
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5530
         Cols            =   6
         Rows            =   10
         FixedCols       =   1
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
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5700
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1260
      TabIndex        =   5
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Libro:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmTipoDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_TIPODOC = 1
Const C_NOMBRE = 2
Const C_DIMINUTIVO = 3
Const C_TIPODOCFIJO = 4
Const C_UPDATE = 5

Dim lTipoLib As Integer

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If lTipoLib <> LIB_OTROS Then
      If ItemData(Cb_TipoLib) = LIB_OTROFULL Then
        MsgBox1 "No es posible modificar la lista de Otros Documentos Full.", vbInformation + vbOKOnly
        Exit Sub
      Else
        MsgBox1 "Sólo es posible modificar la lista de Otros Documentos.", vbInformation + vbOKOnly
        Exit Sub
      End If
   End If
   
   If Val(Grid.TextMatrix(Row, C_ID)) = 0 And Grid.TextMatrix(Row, C_NOMBRE) = "" And Grid.TextMatrix(Row, C_DIMINUTIVO) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   If Val(Grid.TextMatrix(Row, C_TIPODOCFIJO)) <> 0 Then
      MsgBox1 "No se puede borrar este tipo de documento.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   'vemos si está siendo utilizado
   
   Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib= " & lTipoLib & " AND TipoDoc =" & Val(Grid.TextMatrix(Row, C_TIPODOC))
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      MsgBox1 "No se puede borrar este tipo de documento. Está siendo utilizado para algún documento.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
   End If
   
   Call CloseRs(Rs)
   
   If MsgBox1("¿Está seguro que desea borrar este tipo de documento?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_ID, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
         
End Sub

Private Sub Bt_OK_Click()

   If Valida() Then
      Call SaveGrid
      Unload Me
   End If
   
End Sub

Private Sub Cb_TipoLib_Click()
   Dim i As Integer
   Dim Upd As Boolean
   
   If lTipoLib <= 0 Then
      lTipoLib = ItemData(Cb_TipoLib)
      Call LoadAll
      Exit Sub
   End If
      
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_UPDATE) <> "" Then
         Upd = True
         Exit For
      End If
   Next i
         
   If Upd Then
      If MsgBox1("¿Desea grabar los cambios realizados en este libro?", vbYesNo + vbQuestion) = vbYes Then
         Call SaveGrid
      End If
   End If

   lTipoLib = ItemData(Cb_TipoLib)
   
   Call LoadAll
End Sub

Private Sub Form_Load()
   Dim i As Integer

   Call SetUpGrid
   
   lTipoLib = 0
      
   'For i = LIB_COMPRAS To LIB_OTROS
   For i = 1 To UBound(gTipoLibNew)
      Cb_TipoLib.AddItem ReplaceStr(gTipoLibNew(i).Nombre, "Libro de ", "")
      Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).Id 'i
   Next i
   
   Cb_TipoLib.ListIndex = 0

End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_TIPODOC) = 0
   Grid.ColWidth(C_NOMBRE) = 2650
   Grid.ColWidth(C_TIPODOCFIJO) = 0
   Grid.ColWidth(C_UPDATE) = 0
         
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_DIMINUTIVO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_NOMBRE) = "Tipo Documento"
   Grid.TextMatrix(0, C_DIMINUTIVO) = "Dimunutivo"
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
   Next i
   
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
   Grid.TopRow = Grid.FixedRows
   
End Sub

Private Sub LoadAll()
   Dim TipoLib As Integer
   Dim i As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   TipoLib = lTipoLib
   
   If TipoLib <= 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT Id, TipoDoc, Nombre, Diminutivo, TipoDocFijo FROM TipoDocs "
   Q1 = Q1 & " WHERE TipoLib=" & TipoLib & " AND Atributo='ACTIVO'"
   Q1 = Q1 & " ORDER BY TipoDoc"
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_ID) = vFld(Rs("Id"))
      Grid.TextMatrix(i, C_TIPODOC) = vFld(Rs("TipoDoc"))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
      Grid.TextMatrix(i, C_DIMINUTIVO) = vFld(Rs("Diminutivo"))
      Grid.TextMatrix(i, C_TIPODOCFIJO) = IIf(vFld(Rs("TipoDocFijo")), 1, 0)
      
      i = i + 1
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
   Grid.TopRow = Grid.FixedRows
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   
   Action = vbOK
   Value = Trim(Value)

   If Value = "" Then
      MsgBeep vbExclamation
      Action = vbCancel
      Exit Sub
   End If
   
   If Col = C_DIMINUTIVO Then
      
      'vemos si ya existe en otros libros
      
      Q1 = "SELECT Id FROM TipoDocs WHERE Diminutivo='" & Value & "' AND TipoLib <> " & lTipoLib
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If Val(Grid.TextMatrix(Row, C_ID)) <> vFld(Rs("Id")) Then
            MsgBox1 "Este diminutivo ya está definido para otro tipo de documento.", vbExclamation + vbOKOnly
            Action = vbCancel
            Call CloseRs(Rs)
            Exit Sub
         End If
      End If
            
      Call CloseRs(Rs)
      
      'veamos si está definido en la grilla (libro actual)
      
      For i = Grid.FixedRows To Grid.rows - 1
      
         If Grid.TextMatrix(i, C_NOMBRE) = "" And Grid.TextMatrix(i, C_DIMINUTIVO) = "" Then
            Exit For
         End If
         
         If Value = Grid.TextMatrix(i, C_DIMINUTIVO) And i <> Row Then
            MsgBox1 "Este diminutivo ya está definido para otro tipo de documento.", vbExclamation + vbOKOnly
            Action = vbCancel
            Exit Sub
         End If
         
      Next i
      
      If Action = vbOK And Row >= Grid.rows - 2 Then
         Grid.rows = Grid.rows + 1
      End If
      
   End If
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   End If
   
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lTipoLib <> LIB_OTROS Then
      MsgBox1 "Sólo es posible modificar la lista de Otros Documentos.", vbInformation + vbOKOnly
      Exit Sub
   End If
   
   If Row > Grid.FixedRows Then
      
      If Val(Grid.TextMatrix(Row - 1, C_ID)) = 0 And (Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Or Grid.TextMatrix(Row - 1, C_DIMINUTIVO) = "") Then
         MsgBox1 "Línea anterior incompleta.", vbExclamation + vbOKOnly
         Exit Sub
      End If
      
   End If
   
   If Val(Grid.TextMatrix(Row, C_TIPODOCFIJO)) <> 0 Then
      MsgBox1 "Este tipo de documento no se puede modificar.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Select Case Col
   
      Case C_NOMBRE
         Grid.TxBox.MaxLength = 30
         EdType = FEG_Edit

      Case C_DIMINUTIVO
         Grid.TxBox.MaxLength = 3
         EdType = FEG_Edit
         
   End Select
   
End Sub
Private Sub SaveGrid()
   Dim i As Integer
   Dim MaxTipoDoc As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Id As Long
   Dim FldArray(3) As AdvTbAddNew_t
   
   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_NOMBRE) = "" And Grid.TextMatrix(i, C_DIMINUTIVO) = "" Then     'ya terminó la lista de mov.
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
      
         MaxTipoDoc = 1
         
         Q1 = "SELECT Max(TipoDoc) as MaxTipoDoc FROM TipoDocs WHERE TipoLib=" & lTipoLib
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs("MaxTipoDoc")) + 1
         End If
         Call CloseRs(Rs)
         
'         Set Rs = DbMain.OpenRecordset("TipoDocs")
'         Rs.AddNew
'
'         id = vFld(Rs("Id"))
'         Rs.Fields("TipoLib") = lTipoLib
'         Rs.Fields("TipoDoc") = MaxTipoDoc
'         Rs.Fields("Atributo") = "ACTIVO"
'         Rs.Fields("TipoDocFijo") = 0
'
'         Rs.Update
'         Rs.Close
'         Set Rs = Nothing
         

         FldArray(0).FldName = "TipoLib"
         FldArray(0).FldValue = lTipoLib
         FldArray(0).FldIsNum = True
         
         FldArray(1).FldName = "TipoDoc"
         FldArray(1).FldValue = MaxTipoDoc
         FldArray(1).FldIsNum = True
               
         FldArray(2).FldName = "Atributo"
         FldArray(2).FldValue = "ACTIVO"
         FldArray(2).FldIsNum = False
                     
         FldArray(3).FldName = "TipoDocFijo"
         FldArray(3).FldValue = 0
         FldArray(3).FldIsNum = True
         
         Id = AdvTbAddNewMult(DbMain, "TipoDocs", "Id", FldArray)
         
         Grid.TextMatrix(i, C_ID) = Id
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
         
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM TipoDocs WHERE Id = " & Val(Grid.TextMatrix(i, C_ID))
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE Id = " & Val(Grid.TextMatrix(i, C_ID))
         Call DeleteSQL(DbMain, "TipoDocs", Q1)
         
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then 'Update
         Q1 = "UPDATE TipoDocs SET "
         Q1 = Q1 & "  Nombre = '" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
         Q1 = Q1 & ", Diminutivo = '" & Grid.TextMatrix(i, C_DIMINUTIVO) & "'"
         
         Q1 = Q1 & " WHERE Id = " & Val(Grid.TextMatrix(i, C_ID))
         Call ExecSQL(DbMain, Q1)
                  
      End If
      
   Next i

   Call ReadTipoDocs
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   If Grid.Col = C_DIMINUTIVO Then
      Call KeyCod(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
  
End Sub

Private Function Valida() As Boolean
   Dim i As Integer
   Dim Lin As Integer
   
   Valida = False
   
   Lin = 1
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_ID)) = 0 And Grid.TextMatrix(i, C_NOMBRE) = "" And Grid.TextMatrix(i, C_DIMINUTIVO) = "" Then
         Exit For   'terminó la lista
      End If
      
      If Grid.TextMatrix(i, C_NOMBRE) = "" Or Grid.TextMatrix(i, C_DIMINUTIVO) = "" Then
         MsgBox1 "Línea " & Lin & " incompleta. Si desea eliminar el registro, utilice el botón Eliminar.", vbExclamation + vbOKOnly
         Exit Function
      End If
      
      Lin = Lin + 1
                  
   Next i
   
   Valida = True
   
End Function
