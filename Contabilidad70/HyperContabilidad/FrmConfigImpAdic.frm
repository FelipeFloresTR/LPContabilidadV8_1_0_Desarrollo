VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmConfigImpAdic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Impuestos Adicionales Libro de Compras"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   13530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Manual_IEC 
      Caption         =   "Manual IEC"
      Height          =   795
      Left            =   12240
      Picture         =   "FrmConfigImpAdic.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ver Manual Libro Electrónico de Compras"
      Top             =   5100
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Fr_Options 
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   5940
      Visible         =   0   'False
      Width           =   5835
      Begin VB.CheckBox Ch_OcultarImpAdicDescont 
         Caption         =   "Ocultar Impuestos Adicionales descontinuados en Detalle Documento"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.PictureBox Pc_HdCheck 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   12960
      Picture         =   "FrmConfigImpAdic.frx":06B6
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   12
      Top             =   6180
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   13200
      Picture         =   "FrmConfigImpAdic.frx":0A1B
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   6180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Bt_CopyDesdeOtraEmp 
      Caption         =   "Copiar Configuración de otra Empresa..."
      Height          =   435
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copiar la configuración de remuneraciones desde otra empresa"
      Top             =   6060
      Width           =   2955
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5595
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9869
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   1
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   800
      Left            =   12240
      Picture         =   "FrmConfigImpAdic.frx":0A92
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar cuenta seleccionada"
      Top             =   3060
      Width           =   1155
   End
   Begin VB.CommandButton Bt_CopyExcel 
      Caption         =   "Copiar a Excel"
      Height          =   795
      Left            =   12240
      Picture         =   "FrmConfigImpAdic.frx":10F4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Copiar Excel"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton Bt_SelCuenta 
      Caption         =   "Cuentas"
      Height          =   795
      Left            =   12240
      Picture         =   "FrmConfigImpAdic.frx":16A9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cuentas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      Picture         =   "FrmConfigImpAdic.frx":1C44
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Plan de Cuentas"
      Top             =   4080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   12240
      TabIndex        =   7
      Top             =   300
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   12240
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Lb_SelCuentas 
      AutoSize        =   -1  'True
      Caption         =   " Si la tasa está en azul, puede modificarla PARA ESTE DOCUMENTO"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   17
      Top             =   7020
      Visible         =   0   'False
      Width           =   4965
   End
   Begin VB.Label Lb_Config 
      AutoSize        =   -1  'True
      Caption         =   " Si no ha asignado una cuenta, el impuesto adicional no aplica para la empresa"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   7020
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.Label Lb_SelCuentas 
      AutoSize        =   -1  'True
      Caption         =   " Puede seleccionar más de un impuesto adicional a la vez, utilizando la columna de tick."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.Label Lb_Config 
      AutoSize        =   -1  'True
      Caption         =   "NOTAS: Si la Tasa está en blanco, puede ingresarla  PARA ESTA EMPRESA (azul)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   5940
   End
   Begin VB.Label Lb_SelCuentas 
      AutoSize        =   -1  'True
      Caption         =   "NOTAS: Si la Tasa está en blanco, puede ingresala  PARA ESTE DOCUMENTO (verde)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   6285
   End
End
Attribute VB_Name = "FrmConfigImpAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODTIPOVALOR = 0
Const C_IDCUENTA = 1
Const C_CODSII = 2
Const C_TIPOVALOR = 3
Const C_CHECK = 4
Const C_APLICA = 5
Const C_TASA = 6
Const C_ESRECUPERABLE = 7
Const C_CODCUENTA = 8
Const C_CUENTA = 9
Const C_SELCTA = 10
Const C_IDIMPADIC = 11
Const C_TIPODOCAPLICA = 12
Const C_TASAFIJA = 13
Const C_TASAEDITABLE = 14
Const C_UPD = 15

Const NCOLS = C_UPD

Const O_CONFIG = 1
Const O_SELECT = 2

Dim lOper As Integer

Dim lCtasImpAdic(LIBCOMPRAS_NUMOTROSIMP)

Dim lImpAdic() As ImpAdic_t
Dim lRc As Integer

Dim lTipoLib As Integer
Dim lTipoDoc As Integer


Public Sub FConfig(ByVal TipoLib As Integer)
   lOper = O_CONFIG
   lTipoLib = TipoLib
   Me.Show vbModal
   
End Sub
Friend Function FSelect(ByVal TipoLib As Integer, ByVal TipoDoc As Integer, ImpAdic() As ImpAdic_t) As Integer
   Dim i As Integer
   
   lOper = O_SELECT
   lTipoLib = TipoLib
   lTipoDoc = TipoDoc
   Me.Show vbModal
   
   If lRc = vbOK Then
   
      ReDim ImpAdic(UBound(lImpAdic))
      
      For i = 0 To UBound(lImpAdic)
      
         ImpAdic(i) = lImpAdic(i)
       
      Next i
   Else
      ReDim ImpAdic(0)

   End If
   
   
   FSelect = lRc
End Function

Private Sub SetUpGrid()
   Dim wAplica As Integer


   Grid.Cols = NCOLS + 1
   Grid.FixedCols = C_TIPOVALOR + 1
   
   Call FGrSetup(Grid, True)
   
   wAplica = 560
   
   Grid.ColWidth(C_CODTIPOVALOR) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_IDIMPADIC) = 500
   Grid.ColWidth(C_CODSII) = 700
   Grid.ColWidth(C_TIPOVALOR) = 4500
   
   Grid.ColWidth(C_APLICA) = 0
   If lOper = O_CONFIG Then
      Grid.ColWidth(C_APLICA) = wAplica
      Grid.TextMatrix(0, C_APLICA) = "Aplica"
   End If
      
   Grid.ColWidth(C_CHECK) = 0
   If lOper = O_SELECT Then
      Grid.ColWidth(C_CHECK) = 300
'      Grid.Row = 0
'      Grid.Col = C_CHECK
'      Set Grid.CellPicture = Pc_HdCheck
      Call FGrSetPicture(Grid, 0, C_CHECK, Pc_HdCheck, 0)
      Grid.CellPictureAlignment = flexAlignCenterCenter
   End If
   Grid.ColWidth(C_TASA) = 600
   Grid.ColWidth(C_TASAFIJA) = 0
   Grid.ColWidth(C_TASAEDITABLE) = 0
   Grid.ColWidth(C_ESRECUPERABLE) = 700
   Grid.ColWidth(C_CODCUENTA) = 1150
   
   Grid.ColWidth(C_CUENTA) = 3200
   If lOper <> O_CONFIG Then
      Grid.ColWidth(C_CUENTA) = Grid.ColWidth(C_CUENTA) + wAplica
   End If
   
   Grid.ColWidth(C_SELCTA) = IIf(lOper = O_CONFIG, 300, 0)
   Grid.ColWidth(C_TIPODOCAPLICA) = 0
   Grid.ColWidth(C_UPD) = 500
   
   Grid.ColAlignment(C_CODSII) = flexAlignRightCenter
   Grid.ColAlignment(C_TASA) = flexAlignRightCenter
   Grid.ColAlignment(C_ESRECUPERABLE) = flexAlignCenterCenter
   
   
   Grid.TextMatrix(0, C_CODSII) = "Cód. SII"
   Grid.TextMatrix(0, C_TIPOVALOR) = "Tipo de Impuesto Adicional e IVA Retenido"
   Grid.TextMatrix(0, C_TASA) = "Tasa"
   Grid.TextMatrix(0, C_ESRECUPERABLE) = "Recup."
   Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.Col = C_SELCTA
   Grid.Row = 0
   Set Grid.CellPicture = Bt_Cuentas.Picture
   
End Sub

Private Sub Bt_Cancel_Click()

   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_CopyDesdeOtraEmp_Click()
   Dim Frm As FrmCopyPlan

   Set Frm = New FrmCopyPlan
   Call Frm.FCopyConfigImpAdic
   Set Frm = Nothing

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_CopyExcel_Click()

   Call FGr2Clip(Grid, Me.Caption)

End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer

   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_CUENTA) <> "" Then
      If MsgBox1("¿Está seguro que desea eliminar la cuenta asociada a este concepto?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   Grid.Col = C_APLICA
   Grid.Row = Row
   Set Grid.CellPicture = LoadPicture()
   
   Grid.TextMatrix(Row, C_ESRECUPERABLE) = ""
   Grid.TextMatrix(Row, C_IDCUENTA) = ""
   Grid.TextMatrix(Row, C_CODCUENTA) = ""
   Grid.TextMatrix(Row, C_CUENTA) = ""
   Grid.TextMatrix(Row, C_TASA) = ""

   Call FGrModRow(Grid, Row, FGR_D, C_IDIMPADIC, C_UPD)
         
End Sub


Private Sub Bt_Manual_IEC_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_Configuracion_IEC.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual de Configuración del Libro Electrónico de Compras, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual de Configuración del Libro Electrónico de Compras." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub Bt_OK_Click()
   
   If lOper = O_SELECT Then
      If Not GetSelItems Then
         Exit Sub
      End If
   
   ElseIf ValidaConfig Then
      Call SaveAll
   Else
      Exit Sub
   End If
      
   lRc = vbOK
   
   Unload Me
   
   
   
End Sub

Private Sub Bt_SelCuenta_Click()
   Dim Row As Integer
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call Grid_DblClick
End Sub

Private Sub Form_Load()
   Dim i As Integer

   Call SetUpGrid
   
   If lOper = O_SELECT Then
      Me.Caption = "Seleccionar Impuestos Adicionales para agregar a Detalle de Documento"
   Else
      Me.Caption = "Configuración Impuestos Adicionales " & gTipoLib(lTipoLib)
   End If
   
   If lOper = O_SELECT Then
      Bt_Cuentas.visible = False
      Bt_SelCuenta.visible = False
      Bt_CopyExcel.visible = False
      Bt_Del.visible = False
      Bt_Ok.Caption = "Seleccionar"
      Bt_CopyDesdeOtraEmp.visible = False
      For i = 0 To 2
         Lb_SelCuentas(i).visible = True
      Next i
   Else
      For i = 0 To 1
         Lb_Config(i).visible = True
      Next i
      
      Fr_Options.visible = True
   End If
   
   Ch_OcultarImpAdicDescont = IIf(gOcultarImpAdicDescont <> 0, 1, 0)
   
   Call LoadBase

   Call LoadAll

End Sub
Private Sub LoadBase()
   Dim Row As Integer
   Dim i As Integer
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   Row = Grid.rows - 1
   
   For i = 0 To UBound(gTipoValLib)
      If gTipoValLib(i).Nombre <> "" And gTipoValLib(i).TipoLib = lTipoLib And ((gTipoValLib(i).TipoLib = LIB_COMPRAS And gTipoValLib(i).TipoValLib > LIBCOMPRAS_OTROSIMP) Or (gTipoValLib(i).TipoLib = LIB_VENTAS And gTipoValLib(i).TipoValLib > LIBVENTAS_OTROSIMP)) And gTipoValLib(i).Atributo <> "SINUSO" And gTipoValLib(i).Atributo <> "IVAIRREC" And gTipoValLib(i).Atributo <> "IVAACTFIJO" Then
         Grid.rows = Grid.rows + 1
         Row = Row + 1
         Grid.TextMatrix(Row, C_CODTIPOVALOR) = gTipoValLib(i).TipoValLib
         Grid.TextMatrix(Row, C_TIPOVALOR) = IIf(gTipoValLib(i).TitCompleto <> "", gTipoValLib(i).TitCompleto, gTipoValLib(i).Nombre)
         Grid.TextMatrix(Row, C_CODSII) = gTipoValLib(i).CodSIIDTE
         
         'Tasa
         Grid.TextMatrix(Row, C_TASAEDITABLE) = 0
         If gTipoValLib(i).Tasa > 0 Or gTipoValLib(i).TasaFija Then
            Grid.TextMatrix(Row, C_TASAFIJA) = Format(gTipoValLib(i).Tasa, DBLFMT2)
            Grid.TextMatrix(Row, C_TASA) = Grid.TextMatrix(Row, C_TASAFIJA)
         Else
            Call FGrForeColor(Grid, Row, C_TASA, vbBlue)
'            If lOper = O_CONFIG Then
               Grid.TextMatrix(Row, C_TASAEDITABLE) = 1
            End If
'         End If
         
         Grid.TextMatrix(Row, C_TIPODOCAPLICA) = gTipoValLib(i).TipoDoc
         Grid.TextMatrix(Row, C_ESRECUPERABLE) = FmtSiNo(gTipoValLib(i).EsRecuperable, False)
         Grid.TextMatrix(Row, C_SELCTA) = ">>"

      End If
   Next i
   
   Call FGrVRows(Grid, 3)
      
   Grid.FlxGrid.Redraw = True
      
End Sub
Private Sub LoadAll()
   Dim Buf As String
   Dim i As Integer, j As Integer
   Dim IdCuenta As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tipo As Integer
   
   Grid.FlxGrid.Redraw = False
   
   'por si viene de copia de configuración de otra empresa
   For i = Grid.FixedRows To Grid.rows - 1
      Grid.TextMatrix(i, C_IDCUENTA) = ""
      Grid.TextMatrix(i, C_CODCUENTA) = ""
      Grid.TextMatrix(i, C_CUENTA) = ""
      Grid.TextMatrix(i, C_IDIMPADIC) = ""
      Grid.TextMatrix(i, C_UPD) = ""
   Next i
  
   
   Q1 = "SELECT IdTValor, TipoValor.Codigo as TipoValLib, ImpAdic.IdCuenta, Cuentas.Codigo As CodCuenta,"
   Q1 = Q1 & " Cuentas.Descripcion as DescCuenta, ImpAdic.IdImpAdic, ImpAdic.Tasa, ImpAdic.EsRecuperable "
   Q1 = Q1 & " FROM (TipoValor INNER JOIN ImpAdic ON TipoValor.TipoLib = ImpAdic.TipoLib AND TipoValor.Codigo = ImpAdic.TipoValor)"
   Q1 = Q1 & " LEFT JOIN Cuentas ON ImpAdic.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "ImpAdic")
   Q1 = Q1 & " WHERE TipoValor.TipoLib = " & lTipoLib & " AND (Atributo IS NULL OR Atributo <> 'SINUSO')"
   Q1 = Q1 & " AND ImpAdic.IdEmpresa = " & gEmpresa.id & " AND ImpAdic.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY TipoValor.TipoLib, Orden, Valor "

   Set Rs = OpenRs(DbMain, Q1)
   i = Grid.FixedRows
   
   Do While Not Rs.EOF And i <= Grid.rows - 1
      For j = Grid.FixedRows To Grid.rows - 1
         If vFld(Rs("TipoValLib")) = vFmt(Grid.TextMatrix(j, C_CODTIPOVALOR)) Then
            Grid.TextMatrix(j, C_IDCUENTA) = vFld(Rs("IdCuenta"))
            Grid.TextMatrix(j, C_CODCUENTA) = Format(vFld(Rs("CodCuenta")), gFmtCodigoCta)
            Grid.TextMatrix(j, C_CUENTA) = vFld(Rs("DescCuenta"))
            Grid.TextMatrix(j, C_ESRECUPERABLE) = FmtSiNo(vFld(Rs("EsRecuperable")), False)
            If lOper = O_CONFIG Then
               Call FGrSetPicture(Grid, j, C_APLICA, Pc_Check, 0)
            End If
            
            If vFmt(Grid.TextMatrix(j, C_TASAFIJA)) = 0 Then   'puede cambiar la Tasa el cliente
            
               If IsNull(Rs("Tasa")) Or vFld(Rs("Tasa")) < 0 Then
                  Grid.TextMatrix(j, C_TASA) = ""
                  Call FGrForeColor(Grid, j, C_TASA, COLOR_VERDEOSCURO)
                  Grid.TextMatrix(j, C_TASAEDITABLE) = 1
               Else
                  Grid.TextMatrix(j, C_TASA) = Format(vFld(Rs("Tasa")), DBLFMT2)
              End If
               
            End If
            
            Grid.TextMatrix(j, C_IDIMPADIC) = vFld(Rs("IdImpAdic"))
            Exit For
         End If
      Next j
      i = j + 1
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If lOper = O_SELECT Then
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(j, C_TIPOVALOR) = "" Then
            Exit For
         End If
         If Val(Grid.TextMatrix(i, C_IDIMPADIC)) = 0 Then
            Grid.RowHeight(i) = 0
         End If
      Next i
            
   End If
   Call FGrVRows(Grid, 3)
   
   Grid.FlxGrid.Redraw = True
End Sub

Private Sub SaveAll()
   Dim Txt As TextBox
   Dim i As Integer
   Dim Tipo As Integer
   Dim Q1 As String
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      Grid.Row = i
      Grid.Col = C_APLICA

      If Grid.CellPicture <> 0 Then    'sólo se almacenan los que están marcados en la columna APLICA
         
         If Val(Grid.TextMatrix(i, C_IDIMPADIC)) <> 0 Then   'está configurado en la tabla ImpAdic
      
            If Grid.TextMatrix(i, C_UPD) = FGR_U Then
               Q1 = "UPDATE ImpAdic SET IdCuenta = " & Val(Grid.TextMatrix(i, C_IDCUENTA))
               Q1 = Q1 & ", CodCuenta = '" & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA)) & "'"
               Q1 = Q1 & ", EsRecuperable = " & ValSiNo(Grid.TextMatrix(i, C_ESRECUPERABLE), 0)   'si lo deja en blanco es NO
               
               If vFmt(Grid.TextMatrix(i, C_TASAFIJA)) = 0 Then   'si no hay tasa fija, el cliente la puede modificar
                  If Trim(Grid.TextMatrix(i, C_TASA)) = "" Then   'si la deja en blanco, no hay valor y la dejamos en -1
                     Q1 = Q1 & ", Tasa = -1"
                  Else                                            'hay valor, lo guardamos
                     Q1 = Q1 & ", Tasa = " & str(vFmt(Grid.TextMatrix(i, C_TASA)))
                  End If
                 '2991797
               Else
                Q1 = Q1 & ", Tasa = " & str(vFmt(Grid.TextMatrix(i, C_TASA)))
               '2991797
                  
               End If
                     
               Q1 = Q1 & " WHERE IdImpAdic = " & Val(Grid.TextMatrix(i, C_IDIMPADIC))
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)
            End If
            
         Else
         
            Q1 = "INSERT INTO ImpAdic (TipoLib, TipoValor, IdCuenta, CodCuenta, Tasa, EsRecuperable, IdEmpresa, Ano) "
            Q1 = Q1 & " VALUES(" & lTipoLib & "," & Val(Grid.TextMatrix(i, C_CODTIPOVALOR))
            Q1 = Q1 & "," & Val(Grid.TextMatrix(i, C_IDCUENTA))
            Q1 = Q1 & ",'" & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA)) & "'"
            
            If vFmt(Grid.TextMatrix(i, C_TASAFIJA)) = 0 Then   'si no hay tasa fija, el cliente la puede modificar
               If Trim(Grid.TextMatrix(i, C_TASA)) = "" Then   'si la deja en blanco, no hay valor y la dejamos en -1
                  Q1 = Q1 & ", -1"
               Else                                            'hay valor, lo guardamos
                 Q1 = Q1 & ", " & str(vFmt(Grid.TextMatrix(i, C_TASA)))
               End If
            Else
              '2991797
               'Q1 = Q1 & ", -1"
               Q1 = Q1 & ", " & str(vFmt(Grid.TextMatrix(i, C_TASA)))
            '2991797
            
            End If
            
            Q1 = Q1 & ", " & ValSiNo(Grid.TextMatrix(i, C_ESRECUPERABLE), 0)
            Q1 = Q1 & "," & gEmpresa.id
            Q1 = Q1 & "," & gEmpresa.Ano & ")"
            
            Call ExecSQL(DbMain, Q1)
                
         End If
         
      ElseIf Val(Grid.TextMatrix(i, C_IDIMPADIC)) <> 0 Then   'está configurado en la tabla ImpAdic
'         Q1 = "DELETE * FROM ImpAdic "
         Q1 = " WHERE IdImpAdic = " & Val(Grid.TextMatrix(i, C_IDIMPADIC))
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "ImpAdic", Q1)
      End If
                  
   Next i
   
   If lOper = O_CONFIG Then
      'actualizamos
      gOcultarImpAdicDescont = IIf(Ch_OcultarImpAdicDescont <> 0, 1, 0)
      Call UpdParamEmpresa("NOIMPADESC", 0, gOcultarImpAdicDescont)
   End If
  
End Sub

Private Sub Form_Resize()

   Grid.Height = Me.Height - Grid.Top - Bt_CopyDesdeOtraEmp.Height - 1000
   
   If lOper = O_CONFIG Then      'fr_options.visible = true
      Grid.Height = Grid.Height - Fr_Options.Height
      Fr_Options.Top = Grid.Top + Grid.Height + 40
      
      Lb_Config(0).Top = Fr_Options.Top + Fr_Options.Height + 60
      Lb_Config(1).Top = Lb_Config(0).Top + 200
      Lb_SelCuentas(0).Top = Lb_Config(0).Top
      Lb_SelCuentas(1).Top = Lb_SelCuentas(0).Top + 200
      Lb_SelCuentas(2).Top = Lb_SelCuentas(1).Top + 200
   Else
      Lb_Config(0).Top = Grid.Top + Grid.Height + 100
      Lb_Config(1).Top = Lb_Config(0).Top + 200
      Lb_SelCuentas(0).Top = Lb_Config(0).Top
      Lb_SelCuentas(1).Top = Lb_SelCuentas(0).Top + 200
      Lb_SelCuentas(2).Top = Lb_SelCuentas(1).Top + 200
   
   End If
   
   Grid.Width = Me.Width - Grid.Left - Bt_SelCuenta.Width - 360
   Bt_Ok.Left = Grid.Left + Grid.Width + 130
   Bt_Cancel.Left = Bt_Ok.Left
   Bt_SelCuenta.Left = Bt_Ok.Left
   Bt_CopyExcel.Left = Bt_SelCuenta.Left
   Bt_Del.Left = Bt_SelCuenta.Left
   Bt_CopyDesdeOtraEmp.Top = Grid.Top + Grid.Height + 100
   Bt_CopyDesdeOtraEmp.Left = Grid.Left + Grid.Width - Bt_CopyDesdeOtraEmp.Width
   
   Call FGrVRows(Grid, 3)
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Cod As String
   Dim UltimoNivel As Boolean
   Dim NombCta As String, DescCta As String
   Dim IdCuenta As Long
   Dim Tasa As Single
   
   Value = Trim(Value)
   
   If Col = C_TASA Then
      If Value <> "" Then
         
         Tasa = vFmt(Value)
      
         If Tasa >= 0 And Tasa <= 100 Then
            Value = Format(Value, DBLFMT2)
         Else
            Action = vbCancel
            MsgBox1 "Valor inválido.", vbExclamation
            Exit Sub
         End If
      End If
   
      Call FGrModRow(Grid, Row, FGR_U, C_IDIMPADIC, C_UPD)
      
   
   Else   'C_CODCUENTA
      Value = Trim(Value)
      
      Cod = Trim(ReplaceStr(Value, "-", ""))
      If Len(Cod) < Len(VFmtCodigoCta(gFmtCodigoCta)) Then   'asumimos que está usando nombre corto
         NombCta = UCase(Trim(Value))
         Cod = ""
      Else
         NombCta = ""
      End If
      
      IdCuenta = GetIdCuenta(NombCta, Cod, DescCta, UltimoNivel)
      
      If IdCuenta = 0 Then
         MsgBeep vbExclamation
         Action = vbCancel
      
      ElseIf UltimoNivel = False Then
         MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
         Action = vbCancel
      
      Else
         
         Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
         Value = Format(Cod, gFmtCodigoCta)
         Grid.TextMatrix(Row, C_CUENTA) = DescCta
         Call FGrModRow(Grid, Row, FGR_U, C_IDIMPADIC, C_UPD)
         'marcamos la columna APLICA
         Call FGrSetPicture(Grid, Row, C_APLICA, Pc_Check, 0)
         
      End If
   End If
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)

   If Grid.TextMatrix(Row, C_TIPOVALOR) = "" Then
      Exit Sub
   End If
   
   If Col = C_TASA Then
      If Val(Grid.TextMatrix(Row, C_TASAEDITABLE)) <> 0 Then
         EdType = FEG_Edit
      ElseIf lOper = O_SELECT Then
         If Val(Grid.TextMatrix(Row, C_IDCUENTA)) = 0 Then
            MsgBox1 "Debe completar la configuración de la cuenta antes de modificar la tasa y seleccionar este tipo de impuesto adicional.", vbExclamation
         End If
      End If
   ElseIf Col = C_CODCUENTA And lOper = O_CONFIG Then
      EdType = FEG_Edit
   End If
   
   
End Sub

Private Sub Grid_DblClick()
   Dim FrmPlan As FrmPlanCuentas
   Dim DescCta As String
   Dim CodCta As String
   Dim NombCuenta As String
   Dim Row As Integer, Col As Integer
   Dim IdCuenta As Long
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If lOper = O_SELECT And Col <> C_CHECK And Col <> C_TASA And Col <> C_ESRECUPERABLE Then
      If Not ValidImpAdic(Row) Then
         Exit Sub
      End If
      Grid.Row = Row
      Grid.Col = C_CHECK
            
      If Grid.CellPicture = 0 Then
         Call FGrSetPicture(Grid, Row, C_CHECK, Pc_Check, 0)
      Else
         Set Grid.CellPicture = LoadPicture()
      End If
      
      Call Bt_OK_Click
      Exit Sub
   End If

   If Col = C_CHECK Or Col = C_APLICA Then
  
      If Col = C_CHECK Then
         If Not ValidImpAdic(Row) Then
            Exit Sub
         End If
      End If
      
      Grid.Row = Row
      Grid.Col = Col
            
      If Grid.CellPicture = 0 Then
         Call FGrSetPicture(Grid, Row, Col, Pc_Check, 0)
      Else
         Set Grid.CellPicture = LoadPicture()
      End If

      Exit Sub
      
   End If
   
   If Grid.Col <> C_ESRECUPERABLE And Grid.Col <> C_SELCTA And Grid.Col <> C_CUENTA Then
      Exit Sub
   End If
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col = C_ESRECUPERABLE Then
      If UCase(Grid.TextMatrix(Row, Col)) = "SI" Then
         Grid.TextMatrix(Row, Col) = "No"
      Else
         Grid.TextMatrix(Row, Col) = "Si"
      End If
      
      Call FGrModRow(Grid, Row, FGR_U, C_IDIMPADIC, C_UPD)

      Exit Sub
   End If
   
   If lOper <> O_CONFIG Then
      Exit Sub
   End If
   
   'Columna Cuenta
   Set FrmPlan = New FrmPlanCuentas

   If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta, True) = vbOK Then
      If DescCta <> "" Then
         Grid.TextMatrix(Row, C_IDCUENTA) = IdCuenta
         Grid.TextMatrix(Row, C_CODCUENTA) = Format(CodCta, gFmtCodigoCta)
         Grid.TextMatrix(Row, C_CUENTA) = DescCta
         
         'marcamos la columna APLICA
         Call FGrSetPicture(Grid, Row, C_APLICA, Pc_Check, 0)

         Call FGrModRow(Grid, Row, FGR_U, C_IDIMPADIC, C_UPD)
         
     End If

   End If
   Set FrmPlan = Nothing

End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Dim Col As Integer
   
   Col = Grid.Col
   
   If Col = C_CODCUENTA Then
      KeyNum (KeyAscii)
   ElseIf Col = C_TASA Then
      KeyDecPos (KeyAscii)
   End If
   
   'Call KeyUpper(KeyAscii)
End Sub
Private Function ValidImpAdic(ByVal Row As Integer) As Boolean
   
   ValidImpAdic = False
   
   If InStr(Grid.TextMatrix(Row, C_TIPODOCAPLICA), "," & lTipoDoc & ",") = 0 Then   'no aplica al tipo de documento que se indica como parámetro
      MsgBox1 "Este impuesto no se aplica a " & UCase(GetNombreTipoDoc(lTipoLib, lTipoDoc)) & ".", vbExclamation
      Exit Function
   End If
      
   If Val(Grid.TextMatrix(Row, C_IDCUENTA)) = 0 Then
      MsgBox1 "Este impuesto adicional no ha sido configurado con una cuenta contable.", vbExclamation
      Exit Function
   End If
   
   If Trim(Grid.TextMatrix(Row, C_TASA)) = "" Then
      MsgBox1 "No ha sido asignada una tasa para este impuesto adicional.", vbExclamation
      Exit Function
   End If
   
   ValidImpAdic = True
   
End Function
Private Function ValidaConfig() As Boolean
   Dim i As Integer
   
   ValidaConfig = False
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      Grid.Row = i
      Grid.Col = C_APLICA

      If Grid.CellPicture <> 0 Then
        
         If Val(Grid.TextMatrix(i, C_IDCUENTA)) = 0 Then
            MsgBox1 "El tipo de impuesto o IVA Retenido '" & Grid.TextMatrix(i, C_TIPOVALOR) & "' no tiene CUENTA asignada.", vbExclamation
            Grid.Row = i
            Grid.Col = C_CUENTA
            Exit Function
         End If
         
         If Trim(Grid.TextMatrix(i, C_TASA)) = "" Then
            If MsgBox1("Atención:" & vbCrLf & vbCrLf & "El tipo de impuesto o IVA Retenido '" & Grid.TextMatrix(i, C_TIPOVALOR) & "' no tiene TASA asignada." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
               Grid.Row = i
               Grid.Col = C_TASA
               Exit Function
            End If
         End If
  
      End If
      
   Next i
   
   If MsgBox1("ATENCIÓN:" & vbCrLf & vbCrLf & "Sólo se almacenará la configuración de los impuestos o IVA Retenidos marcados en la columna APLICA." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
      Exit Function
   End If
   
   ValidaConfig = True
   
End Function

Private Function GetSelItems() As Boolean
   Dim i As Integer
   Dim j As Integer
   
   GetSelItems = False
   
   ReDim lImpAdic(10)
   j = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      Grid.Row = i
      Grid.Col = C_CHECK

      If Grid.CellPicture <> 0 Then
         
         If Trim(Grid.TextMatrix(i, C_TASA)) = "" Then
            MsgBox1 "Uno de los impuestos adicionales seleccionados no tiene tasa asignada.", vbExclamation
            Exit Function
         End If
         
         If j >= UBound(lImpAdic) Then
            ReDim Preserve lImpAdic(j + 5)
         End If
         
         lImpAdic(j).TipoLib = lTipoLib
         lImpAdic(j).CodTipoValor = Val(Grid.TextMatrix(i, C_CODTIPOVALOR))
         lImpAdic(j).TipoValor = Grid.TextMatrix(i, C_TIPOVALOR)
         lImpAdic(j).IdCuenta = Val(Grid.TextMatrix(i, C_IDCUENTA))
         lImpAdic(j).CodCuenta = Grid.TextMatrix(i, C_CODCUENTA)
         lImpAdic(j).Cuenta = Grid.TextMatrix(i, C_CUENTA)
         lImpAdic(j).Tasa = vFmt(Grid.TextMatrix(i, C_TASA))
         lImpAdic(j).EsRecuperable = ValSiNo(Grid.TextMatrix(i, C_ESRECUPERABLE))
         
         j = j + 1
         
      End If
      
   Next i
         
   GetSelItems = True
   
End Function

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim Col As Integer, Row As Integer
   
   Col = Grid.FlxGrid.MouseCol
   Row = Grid.FlxGrid.MouseRow
  
   If Col = C_CHECK And lOper = O_SELECT Then
      Grid.ToolTipText = "Doble-click para seleccionar impuesto."
   Else
      Grid.ToolTipText = ""
   End If

End Sub

