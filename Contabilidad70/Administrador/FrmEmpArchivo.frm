VERSION 5.00
Begin VB.Form FrmEmpArchivo 
   Caption         =   "Empresas desde Archivo"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_VerFormato 
      Caption         =   "Ver Formato Archivo..."
      Height          =   675
      Left            =   8280
      TabIndex        =   10
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresas Cargadas"
      Height          =   735
      Left            =   1200
      TabIndex        =   7
      Top             =   6240
      Width           =   6855
      Begin VB.Label LblCantidadEmpresas 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Total :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   6915
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   5520
         Picture         =   "FrmEmpArchivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "FrmEmpArchivo.frx":056A
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   480
      Width           =   555
   End
   Begin VB.ListBox Ls_Emp 
      Height          =   4350
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   6855
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Cargar"
      Height          =   315
      Left            =   8280
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8280
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FrmEmpArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lsEmp As ClsCombo

Private Sub Bt_Browse_Click()
FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gAppPath
   If lFileFilter = "" Then
      FrmMain.Cm_ComDlg.Filter = "Archivos CSV(*.csv)|*.csv|Archivos TXT(*.txt)|*.txt|"
   Else
      FrmMain.Cm_ComDlg.Filter = lFileFilter
   End If
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Sub
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   Err.Clear
   
   If lExpectedFName <> "" Then
      If FrmMain.Cm_ComDlg.FileTitle <> lExpectedFName Then
         MsgBox1 "Nombre de archivo inválido." & vbCrLf & vbCrLf & "Nombre esperado: " & lExpectedFName, vbExclamation
         Exit Sub
      End If
   End If
   
   lSelFile = FrmMain.Cm_ComDlg.Filename
   
   Tx_FName = lSelFile
   
   DoEvents
      
End Sub

Public Function FSelect() As Integer

  Me.Show vbModal
  
  FSelect = lRc

End Function

Private Sub Bt_Cancel_Click()
 lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Call ImportFromFile
End Sub

Private Function ImportFromFile() As Boolean
   Dim FName As String
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ImpEnable As Boolean
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim i As Integer, l As Integer
   Dim j As Integer, p As Long, k As Integer
   Dim NUpd As Long
   Dim NIns As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim Aux As String
   Dim vEstado As Integer
   Dim NRecErroneos As Integer, StrNRecErroneos As String
   Dim CampoInvalido As String
   Dim lId As String
   Dim vRut As String
   Dim Sep As String
   Dim vNomCorto As String
   Dim FNameLogImp As String
  
   Dim FldArray(3) As AdvTbAddNew_t
   
    Sep = ";"
   ImportFromFile = False
   
   lFNameLogImp = W.AppPath & "\Importar" & "\Log\ImpEmpresas-" & Format(Now, "yyyymmdd") & ".log"
   
   On Error Resume Next
   
   If Err = cdlCancel Then
      Exit Function
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & Tx_FName, vbExclamation
      Exit Function
   End If

   If Tx_FName = "" Then
      Exit Function
   End If
   Err.Clear
   
   FName = Tx_FName
   
   MousePointer = vbHourglass
   DoEvents
      
   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & FName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
   
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If Err Then
      MsgErr FName
      ImportFromFile = -Err
      Exit Function
   End If
   
   Row = i
   r = 0
   
   Do Until EOF(Fd)
   
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "RUT", vbTextCompare) Then
         GoTo NextRec
      End If
      
      CampoInvalido = ""
         Aux = Trim(NextField2(Buf, p, Sep))
         
       If Not ValidRut((Aux)) Then
          CampoInvalido = CampoInvalido & "," & p
          Call AddLogImp(lFNameLogImp, FName, l, "Rut inválido.")
       Else
         vRut = vFmtCID(Aux, False)
         If vRut = "0" Or vRut = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "RUT inválido")
         End If
       End If
       
         Aux = ""
         Aux = Trim(NextField2(Buf, p, Sep))
         vNomCorto = Aux
         If (vNomCorto = "0" Or vNomCorto = "") Then   'es inválido
            'NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Nombre Corto Obligatorio.")
         End If
         
          If Len(vNomCorto) > 15 Then   'es inválido
            'NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Nombre Corto supera cantidad de caracteres.")
         End If
         
'         Aux = ""
'         Aux = Trim(NextField2(Buf, p, Sep))
'         If (Aux = "") Then    'es inválido

    
      'si no hay errores y la entidad no existe, la insertamos
      
      If CampoInvalido = "" Then
      
         If vRut <> "" And vRut <> "NULO" Then
            IdEnt = 0
            
             Dim x As Integer, nRut As Long, dv As String
               
            x = InStr(UCase(ReplaceStr(Trim(vRut), ".", "")), "-")
                              
            If x = 0 Then
            nRut = vFmtCID(vRut, vRut <> 0)
            Else
            nRut = Val(Left(UCase(ReplaceStr(Trim(vRut), ".", "")), x - 1))
            End If
                  
                  
            Q1 = "SELECT IdEmpresa FROM Empresas WHERE Rut = '" & nRut & "'"
            'Q1 = Q1 & " AND IdEmpresa = " & gEmpresa
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdEnt = vFld(Rs("IdEmpresa"))
               'NombEnt = vFld(Rs("Nombre"))
            End If
            Call CloseRs(Rs)
            
            If IdEnt = 0 Then  'no existe
         
                FldArray(0).FldName = "Rut"
                FldArray(0).FldValue = nRut '& DV_Rut(vRut)
                FldArray(0).FldIsNum = False
                
                FldArray(1).FldName = "NombreCorto"
                FldArray(1).FldValue = ParaSQL(vNomCorto)
                FldArray(1).FldIsNum = False
                
                FldArray(2).FldName = "Import"
                FldArray(2).FldValue = ParaSQL(1)
                FldArray(2).FldIsNum = True

                lId = AdvTbAddNewMult(DbMain, "Empresas", "IdEmpresa", FldArray)
                
                If lId > 0 Then
                Call lsEmp.AddItem(FmtEmprLs(vFmtRut(nRut & DV_Rut(nRut)), vNomCorto), vRut)
                
                 r = r + 1
                End If
            Else
            
            Call AddLogImp(lFNameLogImp, FName, l, "Rut ya Existe")
            NRecErroneos = NRecErroneos + 1
            End If
            
         End If
         
        '3071158
         Call CloseRs(Rs)
         
       '   End If
        '3071158

      Else
         NRecErroneos = NRecErroneos + 1
         
         
      End If
      
NextRec:
   Loop

   Close #Fd
   
   Me.MousePointer = vbDefault
   
   If NRecErroneos = 0 Then
      If r >= 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " Empresas.", vbInformation + vbOKOnly
         LblCantidadEmpresas = r
'      ElseIf r > 1 And Mayor3000reg = False Then
'         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " Empresas.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      ElseIf r > 1 And Mayor3000reg Then
'         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " Empresas, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- No se agregaron Empresas.", vbInformation + vbOKOnly
      End If
   
   Else
      If NRecErroneos > 1 Then
         StrNRecErroneos = "- Se encontraron " & NRecErroneos & " registros con errores en el archivo."
      Else
         StrNRecErroneos = "- Se encontró " & NRecErroneos & " registro con errores en el archivo."
      End If
   
      If r >= 1 Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregó " & r & " Empresas.", vbInformation + vbOKOnly
        LblCantidadEmpresas = r
'      ElseIf r > 1 And Mayor3000reg = False Then
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " Empresas.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      ElseIf r > 1 And Mayor3000reg Then
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- No se agregaron Empresas.", vbInformation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If


   ImportFromFile = True
   
End Function

Private Sub Bt_VerFormato_Click()
   Dim Frm As FrmInforAyudaEmpresas
   
   MousePointer = vbHourglass
   Set Frm = New FrmInforAyudaEmpresas
   Frm.Show vbModal
   Set Frm = Nothing
   MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Set lsEmp = New ClsCombo
   Call lsEmp.SetControl(Ls_Emp)
End Sub
