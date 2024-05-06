VERSION 5.00
Begin VB.Form FrmRevDetDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisar Detalle Documentos"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "FrmRevDetDocs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   1260
      TabIndex        =   3
      Top             =   300
      Width           =   4395
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   5
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Revisar 
      Caption         =   "Revisar"
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   300
      Picture         =   "FrmRevDetDocs.frx":000C
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "FrmRevDetDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Bt_Close_Click()
   Unload Me
End Sub


Private Sub Bt_Revisar_Click()
   Dim Mes As Integer
   Dim Msg As String
   Dim FName As String
   Dim Rc As Long

   Mes = ItemData(Cb_Mes)
   If Mes < 1 Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   Msg = "Este proceso puede tomar algún tiempo dependiendo de la cantidad de documentos ingresados en el sistema."
   
   If MsgBox1(Msg, vbInformation + vbOKCancel + vbDefaultButton2) = vbCancel Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   
   Rc = RevisarDetDocs(Mes, FName)
   If Rc = 0 Then
      MsgBox1 "Proceso finalizado sin errores encontrados en la cuadratura del detalle de los documentos del mes de " & gNomMes(Mes), vbInformation
   
   ElseIf Rc > 0 Then
      MsgBox1 "Se encontraron " & Rc & " documentos con el detalle descuadrado. Revise el archivo " & vbCrLf & vbCrLf & FName, vbExclamation
   
'      If MsgBox1("¿Desea revisar el archivo " & vbCrLf & vbCrLf & FName & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'         Rc = ShellExecute2(Me.hWnd, "open", FName,0, "", SW_SHOW)
'      End If
   
   End If
   
   
   Bt_Revisar.Enabled = True
   MousePointer = vbDefault

End Sub

Private Sub Form_Load()

   Call FillMes(Cb_Mes, GetMesActual())

   Tx_Ano = gEmpresa.Ano

End Sub

Private Function RevisarDetDocs(ByVal Mes As Integer, FName As String) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim Dt1 As Long, Dt2 As Long

   FName = W.AppPath & "\Log\RevDetDocs-" & Format(Now, "yyyymmdd") & ".csv"
   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, Mes, 1), Dt1, Dt2)

   Sep = ";"
   
   Fd = FreeFile
   Err.Clear
   
   Open FName For Output As #Fd
   If Err Then
      MsgErr FName
      RevisarDetDocs = -Err
      Exit Function
   End If

   On Error GoTo 0
   

   'seleccionamos los registros
   Q1 = "SELECT TipoLib, TipoDoc, NumDoc, Entidades.RUT, FEmisionOri, Sum(MovDocumento.Debe) as SumDebe, Sum(MovDocumento.Haber) as SumHaber, Documento.Total "
   Q1 = Q1 & " FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & ")"
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True)
   Q1 = Q1 & " WHERE FEmision BETWEEN " & Dt1 & " AND " & Dt2
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY TipoLib, TipoDoc, NumDoc, RUT, FEmisionOri, Documento.Total "
   Q1 = Q1 & " ORDER BY TipoLib, TipoDoc, NumDoc, RUT, FEmisionOri "
      
   Set Rs = OpenRs(DbMain, Q1)
   
   Buf = "TipoLib" & Sep & "TipoDoc" & Sep & "NumDoc" & Sep & "RUT Entidad" & Sep & "Fecha Emisión" & Sep & "Total"

   Print #Fd, Buf

   Buf = ""
   r = 0
   
   Do While Not Rs.EOF
   
      If vFld(Rs("SumDebe")) <> vFld(Rs("SumHaber")) Then
         Buf = Mid(gTipoLib(vFld(Rs("TipoLib"))), Len("Libro de ")) & Sep & GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & Sep & vFld(Rs("NumDoc")) & Sep & IIf(vFld(Rs("RUT")) <> "", FmtCID(vFld(Rs("RUT"))), "") & Sep & Format(vFld(Rs("FEmisionOri")), EDATEFMT) & Sep & vFld(Rs("Total"))
         Print #Fd, Buf
         r = r + 1
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Close Fd
   
   RevisarDetDocs = r

End Function
