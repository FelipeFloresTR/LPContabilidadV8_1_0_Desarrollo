VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImpEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Empresas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
   Icon            =   "FrmImpEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bt_Descrip 
      Caption         =   "Descripción..."
      Height          =   315
      Left            =   7440
      TabIndex        =   7
      Top             =   1740
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Help 
      Caption         =   "Ayuda (F1)"
      Height          =   315
      Left            =   7440
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Close 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7440
      TabIndex        =   5
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Examinar 
      Caption         =   "Examinar..."
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label La_Import 
         Alignment       =   2  'Center
         Caption         =   "Importando"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   4635
      End
   End
   Begin VB.TextBox tx_Nota 
      Height          =   3435
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   7035
   End
   Begin MSComDlg.CommonDialog CmDlgFile 
      Left            =   8100
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   "Texto separado por tabulaciones (*.txt)|*.txt"
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmImpEmpresa.frx":000C
      Top             =   300
      Width           =   480
   End
End
Attribute VB_Name = "FrmImpEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXREG As Integer = 500

Private lnErr As Integer

Private lFName As String
Private lFmtArray() As FmtImp_t

Private Sub bt_Close_Click()
   Unload Me
End Sub

Private Sub bt_Descrip_Click()
   Dim Frm As FrmFmtImpEnt

   Call FillFmtArray
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FView(Me.Caption, lFName, lFmtArray)
   Set Frm = Nothing

End Sub

Private Sub Bt_Examinar_Click()
   Dim FName As String

   On Error Resume Next

   CmDlgFile.Filename = ""
'   CmDlgFile.InitDir = gPathImport & "\" & gEmpr.Rut ' 6 jun 2017: se agrega el Rut de la empresa
   CmDlgFile.ShowOpen
   CmDlgFile.DialogTitle = "Seleccionar archivo a importar"
   CmDlgFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir

   If Err Then
      If Err <> cdlCancel Then
         MsgErr CmDlgFile.Filename
      End If
      Exit Sub
   End If
   
   Bt_Examinar.Enabled = False
   MousePointer = vbHourglass
   DoEvents
   
   FName = CmDlgFile.Filename
   
   Call ImportEmpr(FName)
   
   CmDlgFile.InitDir = Left(FName, InStrRev(FName, "\"))
   Bt_Examinar.Enabled = True
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_Help_Click()
   Call ShowHelp2(Me, "Import.htm")
End Sub


' 28 may 2018: se agrega detalle de columnas de archivos
Public Function FillFmtArray()
   Dim Cols As String, i As Integer

   ReDim lFmtArray(50)

   lFName = "Empresas.txt"
         
   i = 0
   lFmtArray(i).Campo = "RUT*"
   lFmtArray(i).Formato = "nnnnnnn-v"
   
   i = i + 1
   lFmtArray(i).Campo = "Razón Social*"
   lFmtArray(i).Formato = "Texto (75)"
   
   i = i + 1
   lFmtArray(i).Campo = "Primer Año-Mes*"
   lFmtArray(i).Formato = "aaaa-mm"
   
   i = i + 1
   lFmtArray(i).Campo = "Dirección"
   lFmtArray(i).Formato = "Texto (50)"
   
   i = i + 1
   lFmtArray(i).Campo = "Teléfono"
   lFmtArray(i).Formato = "Texto (50)"
   
   i = i + 1
   lFmtArray(i).Campo = "Fax"
   lFmtArray(i).Formato = "Texto (12)"
   
   i = i + 1
   lFmtArray(i).Campo = "Región"
   lFmtArray(i).Formato = "Texto (30)"
   
   i = i + 1
   lFmtArray(i).Campo = "Ciudad"
   lFmtArray(i).Formato = "Texto (30)"
   
   i = i + 1
   lFmtArray(i).Campo = "Comuna"
   lFmtArray(i).Formato = "Texto (30)"
   
   i = i + 1
   lFmtArray(i).Campo = "email"
   lFmtArray(i).Formato = "Texto (150)"
   
   i = i + 1
   lFmtArray(i).Campo = "URL"
   lFmtArray(i).Formato = "Texto (50)"
   
   i = i + 1
   lFmtArray(i).Campo = "RUT Representante"
   lFmtArray(i).Formato = "nnnnnn-v"
   
   i = i + 1
   lFmtArray(i).Campo = "Nombre Repres."
   lFmtArray(i).Formato = "Texto (50)"
   
   i = i + 1
   lFmtArray(i).Campo = "Giro"
   lFmtArray(i).Formato = "Texto (50)"
      
   i = i + 1
   lFmtArray(i).Campo = "Seguro*"
   lFmtArray(i).Formato = "ISL, Nombre Mutual o código"
         
   i = i + 1
   lFmtArray(i).Campo = "% Seguro"
   lFmtArray(i).Formato = "Real"
   
   i = i + 1
   lFmtArray(i).Campo = "Nombre CCAF o código o blanco"
   lFmtArray(i).Formato = "Texto (2)"
   
   i = i + 1
   lFmtArray(i).Campo = "% CCAF"
   lFmtArray(i).Formato = "Real"
         
   ReDim Preserve lFmtArray(i)
      

End Function


Private Function ImportEmpr(ByVal FName As String) As Long
   Dim Fd As Integer, Buf As String, Q1 As String, Q2 As String, Fld As String, Qry As String
   Dim Rs As Recordset, Dt As Long, AnoMes As Long, idEmpr As Long
   Dim Rc As Long, p As Long, n As Integer, l As Integer, T As Integer, nRut As Long
   Dim Aux As String, nAux As Long, iAux As Integer, dAux As Double, Rut As String
   
   ImportEmpr = -1
      
   If ChkFile(FName) = False Then
      Exit Function
   End If

   tx_Nota = ""
   La_Import = "Importando..."
   La_Import.Visible = True
   DoEvents

   Q1 = "INSERT INTO Empresas ( Rut, RazonSoc, AnoMes, Direccion, Telefono, Fax, CodRegion, CodCiudad, CodComuna, email, Web"
   Q1 = Q1 & ", RutRep, NomRep, Giro, CtaBanco ) VALUES ( "
   
   ' , Seguro/Mutual, PorcSeguro, idCaja, PorcCaja )"

   n = 0
   l = 0
   
   T = LineCount(FName, MAXREG)
   If T = -1 Then
      MsgBox1 "El archivo no debe tener más de " & MAXREG & " registros.", vbExclamation
      Exit Function
   ElseIf T <= 0 Then
      Exit Function
   End If

   
   '**** INICIALIZO BARRA PROGRESIVA
   ProgressBar1.Max = T
   ProgressBar1.Min = 0
   ProgressBar1.Value = 0
   
   On Error Resume Next
   
   Fd = FreeFile
   Open FName For Input As #Fd
   If Err Then
      MsgErr FName
      Exit Function
   End If
   
   Call AddLog("Importando Empresas, archivos: " & FName & ", " & DbGetName(DbMain))
   Call ImpErr("Empresas", "-- Empresas -- [" & FName & "]")
   lnErr = 0
   
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      p = 1
            
      ProgressBar1.Value = l
      DoEvents
      
      Buf = Trim(Buf)
      If Len(Buf) = 0 Or Left(Buf, 1) = "#" Then ' Para los titulos o comentarios
         GoTo NextRec
      End If
      
      Rut = Trim(NextField2(Buf, p)) ' RUT
      If Rut = "" Or (n = 0 And InStr(1, Rut, "rut", vbTextCompare) > 0) Then  ' primera fila con nombres de campos
         GoTo NextRec
      End If
      
      If ValidRut(Rut) = False Then
         Call ImpErr("Empresas", "línea " & l & ": El RUT " & Rut & " es inválido.")
         GoTo NextRec
      End If
      
      nRut = vFmtRut(Rut)
      Qry = "SELECT RazonSoc FROM Empresas WHERE Rut='" & nRut & "'"
      Set Rs = OpenRs(DbMain, Qry)
      If Rs.EOF = False Then
         Call ImpErr("Empresas", "línea " & l & ": El RUT '" & Rut & "' ya está asociado a '" & vFld(Rs("RazonSoc")) & "'.")
         Call CloseRs(Rs)
         GoTo NextRec
      End If
      Call CloseRs(Rs)
      
      Q2 = nRut
      
      Fld = Trim(NextField2(Buf, p))   ' Razón Social
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Fld = Trim(NextField2(Buf, p))   ' aaaa-mm
      If Len(Fld) <> 7 Then
         Call ImpErr("Empresas", "línea " & l & ": Año-Mes inválido '" & Fld & "'.")
         GoTo NextRec
      End If

      nAux = Val(Left(Fld, 4))
      If nAux < Year(Now) - 2 Or nAux > Year(Now) + 2 Then
         Call ImpErr("Empresas", "línea " & l & ": Año inválido '" & nAux & "'.")
         GoTo NextRec
      End If

      iAux = Val(Right(Fld, 2))
      If iAux < 1 Or iAux > 12 Then
         Call ImpErr("Empresas", "línea " & l & ": Mes inválido '" & iAux & "'.")
         GoTo NextRec
      End If

      AnoMes = nAux * 100 + iAux
      Q2 = Q2 & "," & AnoMes

      Fld = Trim(NextField2(Buf, p))   ' Dirección
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Fld = Trim(NextField2(Buf, p))   ' Teléfono
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Fld = Trim(NextField2(Buf, p))   ' Fax
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Fld = Trim(NextField2(Buf, p))   ' Region
      iAux = GetCod(l, Fld, "Regiones", "Regiones", "CodRegion", "Region", 0, "")
      If iAux < 0 Then
         GoTo NextRec
      End If
      Q2 = Q2 & "," & iAux

      Fld = Trim(NextField2(Buf, p))   ' Ciudad
      nAux = GetCod(l, Fld, "Provincias", "Ciudades", "CodCiudad", "Ciudad", iAux, "CodRegion")
      If nAux < 0 Then
         GoTo NextRec
      End If
      Q2 = Q2 & "," & nAux
      iAux = nAux
      
      Fld = Trim(NextField2(Buf, p))   ' Comuna
      nAux = GetCod(l, Fld, "Comunas", "Comunas", "CodComuna", "Comuna", iAux, "CodCiudad")
      If nAux < 0 Then
         GoTo NextRec
      End If
      Q2 = Q2 & "," & nAux
      
      Fld = Trim(NextField2(Buf, p))   ' email
      If Len(Fld) > 0 And ValidEmail(Fld) = False Then
         Call ImpErr("Empresas", "línea " & l & ": email inválido '" & Fld & "'.")
         GoTo NextRec
      End If
      
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"
      
      Fld = Trim(NextField2(Buf, p))   ' URL
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Fld = Trim(NextField2(Buf, p)) ' RUT Representante
      If Len(Fld) > 0 Then
      
         If ValidRut(Rut) = False Then
            Call ImpErr("Empresas", "línea " & l & ": El RUT " & Fld & " es inválido.")
            GoTo NextRec
         End If
   
         Q2 = Q2 & "," & vFmtRut(Fld)
      Else
         Q2 = Q2 & ",NULL"
      End If
      
      Fld = Trim(NextField2(Buf, p))   ' Nombre Representante
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"
      
      Fld = Trim(NextField2(Buf, p))   ' Giro
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"
      
      Fld = Trim(NextField2(Buf, p))   ' Cta Banco
      Q2 = Q2 & ",'" & ParaSQL(Fld) & "'"

      Rc = ExecSQL(DbMain, Q1 & Q2 & " )")

      Qry = "SELECT idEmpresa FROM Empresas WHERE Rut='" & nRut & "'"
      Set Rs = OpenRs(DbMain, Qry)
      If Rs.EOF Then
         Call ImpErr("Empresas", "línea " & l & ": No se pudo crear la empresa " & Rut & ".")
         Call CloseRs(Rs)
         GoTo NextRec
      Else
         idEmpr = vFld(Rs("idEmpresa"))
         Call CloseRs(Rs)
      End If

      Q2 = "INSERT INTO AnoMesEmpr ( idEmpresa, AnoMes, AsigFamPorTramo, RangoSIS, Seguro, PorcSeguro, idCaja, PorcCaja )"
      Q2 = Q2 & " VALUES (" & idEmpr & "," & AnoMes & ",1," & SIS_100oMAS

      Fld = Trim(NextField2(Buf, p))   ' Seguro: ISL o Mutual
      If StrComp(Fld, "ISL", vbTextCompare) = 0 Then
         iAux = SEG_ISL
      Else
         iAux = GetCod(l, Fld, "Mutuales", "Mutual", "$CodPrevired", "Mutual", 0, "")
         If iAux < 0 Then
         
'            Q2 = "DELETE * FROM Empresas WHERE idEmpresa=" & idEmpr
            Call DeleteSQL(DbMain, "Empresas", " WHERE idEmpresa=" & idEmpr)
            
            GoTo NextRec
         End If
      End If
      
      Q2 = Q2 & "," & iAux
      
      Fld = UCase(Trim(NextField2(Buf, p)))   ' % Seguro
      dAux = vFmt(Fld)
      If dAux <= 0 Then
         Call ImpErr("Empresas", "línea " & l & ": El porcentaje del seguro " & Fld & " es inválido.")
         
         Q2 = "DELETE * FROM Empresas WHERE idEmpresa=" & idEmpr
         Call ExecSQL(DbMain, Q2)
            
         GoTo NextRec
      End If
      
      dAux = dAux / 100
      
      Q2 = Q2 & "," & Str0(dAux)
      
      Fld = Trim(NextField2(Buf, p))   ' CCAF
      If Len(Fld) > 0 Then
      
         iAux = GetCod(l, Fld, "CCAF", "Cajas", "$CodPrevired", "Caja", 0, "")
         If iAux < 0 Then
         
            Q2 = "DELETE * FROM Empresas WHERE idEmpresa=" & idEmpr
            Call ExecSQL(DbMain, Q2)
            
            GoTo NextRec
         End If
         
         Q2 = Q2 & "," & iAux
         
         Fld = UCase(Trim(NextField2(Buf, p)))   ' % CCAF
         dAux = vFmt(Fld)
         If dAux <= 0 Then
            Call ImpErr("Empresas", "línea " & l & ": El porcentaje de " & Fld & " es inválido.")
         
            Q2 = "DELETE * FROM Empresas WHERE idEmpresa=" & idEmpr
            Call ExecSQL(DbMain, Q2)
            
            GoTo NextRec
         End If
         
         dAux = dAux / 100
      
         Q2 = Q2 & "," & Str0(dAux)
      Else
         Q2 = Q2 & ", NULL, NULL"
      End If
      
      Rc = ExecSQL(DbMain, Q2 & " )")
      
      n = n + 1

NextRec:
   Loop
   
   Close #Fd

   La_Import = ""

   Call ImpErr("Empresas", "Se importaron " & n & " empresas.")

End Function



Private Function ChkFile(ByVal FName As String) As Boolean
   Dim Dt As Double

   ChkFile = False
   
   If Not ExistFile(FName) Then
      MsgBox1 "Archivo " & FName & " no encontrado.", vbExclamation
      Exit Function
   End If
   
   On Error Resume Next
   Dt = FileDateTime(FName)
   If Dt = 0 Then
      MsgErr FName
      Exit Function
   ElseIf Dt < Now - TimeSerial(5, 0, 0) Then
   
      If MsgBox1("El archivo " & FName & " es muy antiguo. " & Format(Dt, "d mmm yyyy hh:nn") & vbCrLf & "¿ Desea continuar?", vbYesNo Or vbExclamation Or vbDefaultButton2) <> vbYes Then
         Exit Function
      End If
   End If
      
   ChkFile = True
   
End Function

Private Sub ImpErr(ByVal Tabla As String, ByVal Msg As String)
   
   Call AddTxt(tx_Nota, Msg)  ' , True)
   Call AddLogErr(Tabla, Msg)
   
   lnErr = lnErr + 1
   
End Sub

Private Sub AddLogErr(ByVal Tabla As String, ByVal Msg As String)
   Dim Fd As Long
   Dim eErr As Long, eDesc As String, eDLL As Long
      
   On Error Resume Next
   Fd = FreeFile
   Open w.AppPath & "\Log\" & Tabla & ".log" For Append Access Write As #Fd

   Print #Fd, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & GetComputerName() & vbTab & gUsuario.Nombre & vbTab & Msg
   
   Close #Fd
   
End Sub

Private Function GetCod(ByVal l As Integer, ByVal Valor As String, ByVal Txt As String, ByVal Tabla As String, ByVal CodCampo As String, ByVal Campo As String, ByVal idPadre As Integer, ByVal CodPadre As String)
   Dim Rs As Recordset, Q1 As String, bStr As Boolean

   If Valor = "" Then
      Exit Function
   End If

   If Left(CodCampo, 1) = "$" Then
      bStr = True
      CodCampo = Mid(CodCampo, 2)
   End If

   If IsNumeric(Valor) Then
      
      Q1 = "SELECT " & CodCampo & " FROM " & Tabla
      
      If bStr Then
         Q1 = Q1 & " WHERE " & CodCampo & "='" & Valor & "'"
      Else
         Q1 = Q1 & " WHERE " & CodCampo & "=" & Valor
      End If
      
      If idPadre > 0 Then
         Q1 = Q1 & " And " & CodPadre & "=" & idPadre
      End If
      
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF Then
         Call ImpErr("Empresas", "línea " & l & ": El código '" & Valor & "' no existe en '" & Txt & "'.")
         GetCod = -1
      Else
         GetCod = Val(Valor)
      End If
      Call CloseRs(Rs)
      
   Else
      Q1 = "SELECT " & CodCampo & " FROM " & Tabla & " WHERE " & Campo & "='" & ParaSQL(Valor) & "'"
      
      If idPadre > 0 Then
         Q1 = Q1 & " And " & CodPadre & "=" & idPadre
      End If
      
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF Then
         Call ImpErr("Empresas", "línea " & l & ": El nombre '" & Valor & "' no existe en '" & Txt & "'.")
         GetCod = -2
      Else
         GetCod = vFld(Rs(CodCampo))
      End If
      Call CloseRs(Rs)
   End If

End Function
