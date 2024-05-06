VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConfigInformes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuraciones para Informes"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18270
   Icon            =   "FrmConfigInformes.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   18270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Documentos ODF"
      Height          =   780
      Left            =   9240
      TabIndex        =   52
      Top             =   2280
      Width           =   7695
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   5640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   240
         Width           =   1545
      End
      Begin VB.CheckBox Ch_DocAnalitico 
         Caption         =   "Incluir en Informe Analítico"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   8
         Left            =   5040
         TabIndex        =   55
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Frame Fr_PieFirma 
      Caption         =   "Pie de Firma para Comprobantes (Impresión)"
      Height          =   1905
      Left            =   9240
      TabIndex        =   43
      Top             =   3240
      Width           =   7695
      Begin VB.TextBox Tx_TextoMembrete2 
         Height          =   375
         Left            =   5160
         TabIndex        =   49
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Tx_TextoMembrete1 
         Height          =   375
         Left            =   5160
         TabIndex        =   47
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Tx_Membrete2 
         Height          =   375
         Left            =   1200
         TabIndex        =   48
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Tx_Membrete1 
         Height          =   375
         Left            =   1200
         TabIndex        =   46
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Texto 2"
         Height          =   375
         Left            =   4440
         TabIndex        =   51
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Texto 1"
         Height          =   255
         Left            =   4440
         TabIndex        =   50
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Titulo del Membrete 2"
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Titulo del Membrete 1"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame FrameEgreso 
      Caption         =   "Egreso (Activo Realizable o Prestacion de Servicio)"
      Height          =   855
      Left            =   1260
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   7755
      Begin VB.TextBox Txt_EgresoServicio 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txt_EgresoExistencia 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   37
         Text            =   "100"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "%"
         Height          =   195
         Left            =   5640
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Servicios"
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Existencias o Insumos"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         Height          =   195
         Left            =   2880
         TabIndex        =   38
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.Frame Fr_NumReg 
      Caption         =   "Registros por página de reporte"
      Height          =   855
      Left            =   9240
      TabIndex        =   33
      Top             =   5280
      Width           =   7755
      Begin VB.TextBox Tx_NumReg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1920
         TabIndex        =   12
         Text            =   "1.000"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "100 - 1.000"
         Height          =   195
         Left            =   2940
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad de registros:"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   2
      Left            =   300
      Picture         =   "FrmConfigInformes.frx":000C
      ScaleHeight     =   705
      ScaleWidth      =   675
      TabIndex        =   32
      Top             =   360
      Width           =   675
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   17055
      TabIndex        =   14
      Top             =   720
      Width           =   1035
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   17055
      TabIndex        =   13
      Top             =   360
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1905
      Index           =   1
      Left            =   1260
      TabIndex        =   29
      Top             =   3240
      Width           =   7755
      Begin VB.ListBox Ls_Opt 
         Height          =   960
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   315
         Width           =   7245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notas "
      Height          =   1890
      Index           =   0
      Left            =   1260
      TabIndex        =   26
      Top             =   270
      Width           =   15630
      Begin VB.CheckBox Ch_NotaEspBal 
         Caption         =   "Incluir en Balances"
         Height          =   255
         Left            =   9660
         TabIndex        =   31
         Top             =   315
         Width           =   1815
      End
      Begin VB.CheckBox Ch_Art100Lib 
         Caption         =   "Incluir en Libros"
         Height          =   255
         Left            =   3780
         TabIndex        =   30
         Top             =   270
         Width           =   1455
      End
      Begin VB.CheckBox Ch_NotaEspInfo 
         Caption         =   "Incluir en otros Informes"
         Height          =   255
         Left            =   13200
         TabIndex        =   5
         Top             =   315
         Width           =   1995
      End
      Begin VB.CheckBox Ch_Art100Info 
         Caption         =   "Incluir en otros Informes"
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   270
         Width           =   1995
      End
      Begin VB.CheckBox Ch_NotaEspLib 
         Caption         =   "Incluir en Libros"
         Height          =   255
         Left            =   11580
         TabIndex        =   4
         Top             =   315
         Width           =   1515
      End
      Begin VB.CheckBox Ch_Art100Bal 
         Caption         =   "Incluir en Balances"
         Height          =   255
         Left            =   1860
         TabIndex        =   1
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox Tx_NotaEsp 
         Height          =   1215
         Left            =   7980
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   7230
      End
      Begin VB.TextBox Tx_Art100 
         Height          =   1215
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   555
         Width           =   7230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota especial:"
         Height          =   195
         Index           =   1
         Left            =   7980
         TabIndex        =   28
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota artículo 100:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   315
         Width           =   1290
      End
   End
   Begin VB.Frame Fr_Color 
      Caption         =   "Color por Nivel"
      Height          =   780
      Left            =   1260
      TabIndex        =   15
      Top             =   2280
      Width           =   7740
      Begin VB.CommandButton Bt_Color 
         Height          =   375
         Index           =   1
         Left            =   1215
         Picture         =   "FrmConfigInformes.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   285
         Width           =   375
      End
      Begin VB.CommandButton Bt_Color 
         Height          =   375
         Index           =   2
         Left            =   2745
         Picture         =   "FrmConfigInformes.frx":0AA7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Bt_Color 
         Height          =   375
         Index           =   3
         Left            =   4230
         Picture         =   "FrmConfigInformes.frx":0EDC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Bt_Color 
         Height          =   375
         Index           =   4
         Left            =   5715
         Picture         =   "FrmConfigInformes.frx":1311
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Bt_Color 
         Height          =   375
         Index           =   5
         Left            =   7245
         Picture         =   "FrmConfigInformes.frx":1746
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel 1:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   25
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel 2:"
         Height          =   255
         Index           =   5
         Left            =   1740
         TabIndex        =   24
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel 3:"
         Height          =   255
         Index           =   2
         Left            =   3195
         TabIndex        =   23
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel 4:"
         Height          =   255
         Index           =   3
         Left            =   4725
         TabIndex        =   22
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel 5:"
         Height          =   255
         Index           =   4
         Left            =   6255
         TabIndex        =   21
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Lb_Color 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   795
         TabIndex        =   20
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Lb_Color 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2325
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lb_Color 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3825
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lb_Color 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   5310
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lb_Color 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   6825
         TabIndex        =   16
         Top             =   345
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog Cm_ComDlg 
      Left            =   8505
      Top             =   3555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmConfigInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lColores(MAX_NIVELES) As Long

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Color_Click(Index As Integer)
   Cm_ComDlg.CancelError = True
   On Error Resume Next
   Cm_ComDlg.DialogTitle = "Seleccionar Color"
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If
      
   Cm_ComDlg.ShowColor
   If Cm_ComDlg.Color <> 0 Then
      lColores(Index) = Cm_ComDlg.Color
   End If
   Lb_Color(Index).BackColor = lColores(Index)

End Sub

Private Sub Bt_OK_Click()

   Call SaveNotas
   Call SaveColores
   Call SaveOpciones
   Call SaveParametrosODF
   
   '2860036
    If valida() = True Then
        Call SaveMembrete
    End If
   'fin 2860036
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   'NOTAS
   Call LoadAllNotas
   
   'COLORES
   Call LoadAllColores
   
   'OPCIONES
   Call FillList(gEmpresa.Opciones)
   
   '2860036
    Call LoadAllMembrete
   'fin 2860036
   
   'Call CargaPocentaje
   
   If gDbType = SQL_ACCESS Then
      Fr_NumReg.visible = False
      Me.Height = Me.Height - 800
   Else
      Tx_NumReg = Format(gPageNumReg, NUMFMT)
      
   End If
   
'   For i = 1 To UBound(gTratamiento)
'      Cb_Tratamiento.AddItem ReplaceStr(gTratamiento(i).Nombre, "Libro de ", "")
'      Cb_Tratamiento.ItemData(Cb_Tratamiento.NewIndex) = gTratamiento(i).Id 'i
'   Next i
'   Cb_Tratamiento.ListIndex = 1
   
   For i = 1 To MAX_ESTADODOC
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
      
   Call ParametrosODF
   Cb_Estado.ListIndex = 1
  
End Sub
Private Sub ParametrosODF()
Dim Q1 As String
Dim Rs As Recordset

   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='INFANAODF' "
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = True Then
        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'INFANAODF', 1, '0')"
        Call ExecSQL(DbMain, Q1)
   Else
       Ch_DocAnalitico = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)
   
'   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='TRATAMODF' "
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Rs.EOF = True Then
'        Q1 = "INSERT INTO Param "
'        Q1 = Q1 & " (Tipo, Codigo, Valor)"
'        Q1 = Q1 & " VALUES( 'TRATAMODF', 1, '1')"
'        Call ExecSQL(DbMain, Q1)
'   Else
'        Call SelItem(Cb_Tratamiento, vFld(Rs("Valor")))
'   End If
'   Call CloseRs(Rs)
   
   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='ESTADOODF' "
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = True Then
        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'ESTADOODF', 1, '1')"
        Call ExecSQL(DbMain, Q1)
   Else
        Call SelItem(Cb_Estado, vFld(Rs("Valor")))
   End If
   Call CloseRs(Rs)


End Sub
Private Sub SaveParametrosODF()
Dim Q1 As String

    Q1 = "UPDATE Param SET "
    Q1 = Q1 & " Valor = '" & IIf(Ch_DocAnalitico <> 0, 1, 0) & "'"
    Q1 = Q1 & " WHERE Tipo='INFANAODF' "
    Call ExecSQL(DbMain, Q1)
    
    Q1 = "UPDATE Param SET "
    Q1 = Q1 & " Valor = '" & ItemData(Cb_Estado) & "'"
    Q1 = Q1 & " WHERE Tipo='ESTADOODF' "
    Call ExecSQL(DbMain, Q1)
    
End Sub
Private Sub LoadAllNotas()
   Tx_Art100 = gNotaArt100.TxtNota
   Ch_Art100Bal = Abs(gNotaArt100.IncluirBal)
   Ch_Art100Lib = Abs(gNotaArt100.IncluirLib)
   Ch_Art100Info = Abs(gNotaArt100.IncluirInfo)
   
   Tx_NotaEsp = gNotaEspecial.TxtNota
   Ch_NotaEspBal = Abs(gNotaEspecial.IncluirBal)
   Ch_NotaEspLib = Abs(gNotaEspecial.IncluirLib)
   Ch_NotaEspInfo = Abs(gNotaEspecial.IncluirInfo)


End Sub
Private Sub LoadAllColores()
   Dim i As Integer
   
   For i = 1 To MAX_NIVELES
      Lb_Color(i).BackColor = gColores(i)
      lColores(i) = gColores(i)
   Next i
   
   Call SetupPriv
End Sub
Private Sub SaveNotas()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Incluir As Integer
   
   Incluir = IIf(Ch_Art100Lib <> 0, C_INCNOTALIB, 0) Or IIf(Ch_Art100Bal <> 0, C_INCNOTABAL, 0)
   
   Q1 = "SELECT Nota FROM Notas WHERE Tipo='ART100'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
      'existe, actualizamos
      
      Q1 = "UPDATE Notas SET Nota = '" & ParaSQL(Tx_Art100) & "'"
      Q1 = Q1 & ", Incluir = " & Incluir
      Q1 = Q1 & ", IncluirInfo = " & CInt(Ch_Art100Info <> 0)
      Q1 = Q1 & "  WHERE Tipo = 'ART100'"
      Q1 = Q1 & "  AND IdEmpresa = " & gEmpresa.id
      
   Else
      Q1 = "INSERT INTO Notas "
      Q1 = Q1 & " (Tipo, Nota, Incluir, IncluirInfo, IdEmpresa) "
      Q1 = Q1 & " VALUES ( 'ART100','" & ParaSQL(Tx_Art100) & "', " & Incluir & ", " & CInt(Ch_Art100Info <> 0) & "," & gEmpresa.id & ")"
   End If
   
   Call CloseRs(Rs)
   
   Call ExecSQL(DbMain, Q1)
   
   Incluir = IIf(Ch_NotaEspLib <> 0, C_INCNOTALIB, 0) Or IIf(Ch_NotaEspBal <> 0, C_INCNOTABAL, 0)

   Q1 = "SELECT Nota FROM Notas WHERE Tipo='NOTAESP'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
      'existe, actualizamos
      
      Q1 = "UPDATE Notas SET Nota = '" & ParaSQL(Tx_NotaEsp) & "'"
      Q1 = Q1 & ", Incluir = " & Incluir
      Q1 = Q1 & ", IncluirInfo = " & CInt(Ch_NotaEspInfo <> 0)
      Q1 = Q1 & "  WHERE Tipo = 'NOTAESP'"
      Q1 = Q1 & "  AND IdEmpresa = " & gEmpresa.id
     
   Else
      Q1 = "INSERT INTO Notas "
      Q1 = Q1 & " (IdEmpresa, Tipo, Nota, Incluir, IncluirInfo) "
      Q1 = Q1 & " VALUES ( " & gEmpresa.id & ", 'NOTAESP','" & ParaSQL(Tx_NotaEsp) & "', " & Incluir & ", " & CInt(Ch_NotaEspInfo <> 0) & ")"
   End If
   
   Call CloseRs(Rs)
   
   Call ExecSQL(DbMain, Q1)
   
   gNotaArt100.TxtNota = Tx_Art100
   gNotaArt100.IncluirBal = (Ch_Art100Bal <> 0)
   gNotaArt100.IncluirLib = (Ch_Art100Lib <> 0)
   gNotaArt100.IncluirInfo = (Ch_Art100Info <> 0)
   
   gNotaEspecial.TxtNota = Tx_NotaEsp
   gNotaEspecial.IncluirBal = (Ch_NotaEspBal <> 0)
   gNotaEspecial.IncluirLib = (Ch_NotaEspLib <> 0)
   gNotaEspecial.IncluirInfo = (Ch_NotaEspInfo <> 0)
   
   Call SetPrtNotas
   
End Sub
Private Sub SaveColores()
   Dim Rs As Recordset
   Dim i As Integer

   For i = 1 To MAX_NIVELES
      gColores(i) = lColores(i)
   Next i
   
   Set Rs = OpenRs(DbMain, "SELECT Nivel FROM Colores WHERE IdEmpresa = " & gEmpresa.id)
   
   If Rs.EOF = False Then  'están
      For i = 1 To MAX_NIVELES
         Call ExecSQL(DbMain, "UPDATE Colores SET Color=" & lColores(i) & " WHERE Nivel=" & i & " AND IdEmpresa = " & gEmpresa.id)
      Next i
   
   Else
      For i = 1 To MAX_NIVELES
         Call ExecSQL(DbMain, "INSERT INTO Colores (Nivel, Color, IdEmpresa) VALUES(" & i & "," & lColores(i) & "," & gEmpresa.id & ")")
      Next i
   
   End If
   
   Call CloseRs(Rs)
End Sub
Private Sub SaveOpciones()
   Dim Opt As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   
   Opt = 0
   For i = 0 To Ls_Opt.ListCount - 1
      If Ls_Opt.Selected(i) Then
         Opt = Opt Or Ls_Opt.ItemData(i)
      End If
   Next i
   
   Q1 = "UPDATE Empresa SET Opciones=" & Opt & " WHERE Id = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)
   
'   Q1 = "UPDATE Empresa SET Opciones=" & Opt & ", PorcExisteOServ = " & Me.Txt_EgresoExistencia.Text & " WHERE Id = " & gEmpresa.id
'   Call ExecSQL(DbMain, Q1)
   
   gEmpresa.Opciones = Opt
   
   Call ResetPrtBas(gPrtLibros)
   Call ResetPrtBas(gPrtReportes)
   
   gPageNumReg = vFmt(Tx_NumReg)
   Call SetIniString(gIniFile, "Reportes", "PageNumReg", gPageNumReg)
   

End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   
End Function
Private Sub FillList(Opciones As Long)
   Dim i As Integer
   
   Ls_Opt.Clear
   Ls_Opt.AddItem "Actualizar automáticamente folios usados en la impresión con papel foliado "
   Ls_Opt.ItemData(Ls_Opt.NewIndex) = OPT_ACTUSADO
   Ls_Opt.AddItem "No imprimir fecha en los reportes, sean estos oficiales o no "
   Ls_Opt.ItemData(Ls_Opt.NewIndex) = OPT_NOPRTFECHA
   
   For i = 0 To Ls_Opt.ListCount - 1
      Ls_Opt.Selected(i) = ((Opciones And Ls_Opt.ItemData(i)) <> 0)
      Ls_Opt.Selected(i) = ((Opciones And Ls_Opt.ItemData(i)) <> 0)
   Next i
   
End Sub

Private Sub CargaPocentaje()
Dim Q1 As String
Dim Rs As Recordset

'   Q1 = "Select PorcExisteOServ From empresa WHERE Id = " & gEmpresa.id
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'      Me.Txt_EgresoExistencia.Text = vFld(Rs("PorcExisteOServ"))
'      Me.Txt_EgresoServicio.Text = 100 - Me.Txt_EgresoExistencia.Text
'   Else
'      Me.Txt_EgresoExistencia.Text = 100
'      Me.Txt_EgresoServicio.Text = 0
'   End If
'   Call CloseRs(Rs)
End Sub

Private Sub Tx_NumReg_KeyPress(KeyAscii As Integer)

   Call KeyNumPos(KeyAscii)
   
End Sub

Private Sub Tx_NumReg_LostFocus()

   Tx_NumReg = Format(vFmt(Tx_NumReg), NUMFMT)
   
End Sub

Private Sub Tx_NumReg_Validate(Cancel As Boolean)

   If vFmt(Tx_NumReg) < 100 Or vFmt(Tx_NumReg) > 1000 Then
      MsgBox1 "Valor inválido.", vbExclamation
      Cancel = True
   Else
      Cancel = False
   End If
End Sub

Private Sub Txt_EgresoExistencia_Change()
If Me.Txt_EgresoExistencia.Text > 100 Then
    Me.Txt_EgresoExistencia.Text = 100
    Me.Txt_EgresoExistencia.Text = 0
Else
    Me.Txt_EgresoServicio.Text = Abs(100 - CInt(Me.Txt_EgresoExistencia.Text))
End If
End Sub

Private Sub Txt_EgresoExistencia_KeyPress(KeyAscii As Integer)

Call KeyNumPos(KeyAscii)
If Me.Txt_EgresoExistencia.Text = "" Then
Txt_EgresoExistencia.Text = 0
End If


End Sub


'pipe 2860036
Private Sub SaveMembrete()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Incluir As Integer


   Q1 = "SELECT TituloMembrete1, TituloMembrete2, Texto1, Texto2 FROM Membrete "
   Q1 = Q1 & " where IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then
      'existe, actualizamos

      Q1 = "UPDATE Membrete SET TituloMembrete1 = '" & ParaSQL(Tx_Membrete1) & "'"
      Q1 = Q1 & ", TituloMembrete2 = '" & ParaSQL(Tx_Membrete2) & "'"
      Q1 = Q1 & ", Texto1 = '" & ParaSQL(Tx_TextoMembrete1) & "'"
      Q1 = Q1 & ", Texto2 = '" & ParaSQL(Tx_TextoMembrete2) & "'"
      Q1 = Q1 & "  WHERE IdEmpresa = " & gEmpresa.id

   Else
      Q1 = "INSERT INTO Membrete "
      Q1 = Q1 & " (TituloMembrete1, TituloMembrete2, Texto1, Texto2, IdEmpresa) "
      Q1 = Q1 & " VALUES ( '" & ParaSQL(Tx_Membrete1) & "','" & ParaSQL(Tx_Membrete2) & "','" & ParaSQL(Tx_TextoMembrete1) & "', '" & ParaSQL(Tx_TextoMembrete2) & "'," & gEmpresa.id & ")"
   End If

   Call CloseRs(Rs)

   Call ExecSQL(DbMain, Q1)
   
   gMembrete.TxtTitMembrete1 = Tx_Membrete1
   gMembrete.TxtTitMembrete2 = Tx_Membrete2
   gMembrete.TxtTexto1 = Tx_TextoMembrete1
   gMembrete.TxtTexto2 = Tx_TextoMembrete2



End Sub
' fin 2860036

'2860036
Private Sub LoadAllMembrete()
Dim Q1 As String
   Dim Rs As Recordset
   Dim Incluir As Integer


   Q1 = "SELECT TituloMembrete1, TituloMembrete2, Texto1, Texto2 FROM Membrete "
   Q1 = Q1 & " where IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then
   Tx_Membrete1 = vFld(Rs("TituloMembrete1"))
   Tx_Membrete2 = vFld(Rs("TituloMembrete2"))
   Tx_TextoMembrete1 = vFld(Rs("Texto1"))
   Tx_TextoMembrete2 = vFld(Rs("Texto2"))
   
   End If
 
    Call CloseRs(Rs)
End Sub
'fin 2860036

'2860036
Private Function valida() As Boolean

   valida = False
       
   If Len(Trim(Tx_Membrete1)) > 0 Then
    If Len(Trim(Tx_TextoMembrete1)) = 0 Then
        Call MsgBox1("Falta ingresar Text 1.", vbOKOnly + vbExclamation)
        Exit Function
    End If
   End If
   
   If Len(Trim(Tx_Membrete2)) > 0 Then
    If Len(Trim(Tx_TextoMembrete2)) = 0 Then
        Call MsgBox1("Falta ingresar Text 2.", vbOKOnly + vbExclamation)
        Exit Function
    End If
   End If
   
         
     valida = True
End Function
'fin 2860036
