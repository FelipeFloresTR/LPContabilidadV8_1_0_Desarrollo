VERSION 5.00
Begin VB.Form FrmEmpHR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas en Hyper Renta"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "FrmEmpHR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   240
      Picture         =   "FrmEmpHR.frx":000C
      ScaleHeight     =   780
      ScaleWidth      =   765
      TabIndex        =   4
      Top             =   360
      Width           =   765
   End
   Begin VB.TextBox tx_HR 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Ubicación de la base de HR"
      Top             =   5640
      Width           =   5175
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Seleccionar"
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox Ls_Emp 
      Height          =   5325
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "FrmEmpHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'password LexContab.mdb
Private Const PASSW_LEXCONT = "Fw#420!&+"
Private Const PASSW_PREFIX = "Fw#42+"   'prefijo password empresa (sigue RUT sin puntos, ni guión, ni dígito verificador

Private lConnStr As String
Private lsEmp As ClsCombo
Private lHrConnStr As String
Private lDbPath As String

Dim lHrDb As Database

Public Function FillEmpHR() As Boolean
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim Rc As Long
   Dim i As Integer
   
   FillEmpHR = False
   Rc = 0

   On Error Resume Next
'   Set Conn = New ADODB.Connection
   lDbPath = gHRPath & "\PAR\BD_HR_admin.mdb"
   tx_HR = GetAbsPath(lDbPath, FrmMain.Drive1)
   If ExistFile(lDbPath) = False Then
      MsgBox1 "No se encuentra la base de HR en" & vbCrLf & lDbPath, vbExclamation
      Exit Function
   End If
      
'   lHrConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & lDbPath & ";Jet OLEDB:Database Password=" & "20" & "080" & "3hr" & ";"
   
   Set lHrDb = OpenDatabase(lDbPath, False, False, ";PWD=" & "20" & "080" & "3hr" & ";")
   
'   Call Conn.Open(lHrConnStr)
   If Err Then
      MsgBox "Error H" & Hex(Err) & ", " & Error & NL & lDbPath, vbExclamation
      Exit Function
   End If

   Q1 = "SELECT NC_NomCorto, NC_Rut "
   Q1 = Q1 & " FROM Adm_NContrib"
   Q1 = Q1 & " ORDER BY NC_NomCorto"
   
   Set Rs = OpenRsDao(lHrDb, Q1)
   
   If Rs Is Nothing Then
      Exit Function
   End If
   
   i = 0
   
   Do While Not Rs.EOF
   
'      Call lsEmp.AddItem(vFldDao(Rs("NC_NomCorto")), vFldDao(Rs("NC_Rut")))
      Call lsEmp.AddItem(FmtEmprLs(vFmtRut(vFldDao(Rs("NC_Rut"))), vFldDao(Rs("NC_NomCorto"))), vFldDao(Rs("NC_Rut")))
            
      i = i + 1
      
      Rs.MoveNext
   Loop
   
   If i > 0 Then
      Rc = i
      FillEmpHR = True
   End If
      
   Call CloseRs(Rs)
   
End Function

Public Function Import() As Boolean
   Dim i As Integer, Q1 As String
   Dim Rc As Long
   Dim IdContrib As Long
   Dim Rs As dao.Recordset
   
   i = lsEmp.ListIndex
   Import = False
   
   If i < 0 Then
      Exit Function
   End If

   If lHrDb Is Nothing Then
      Exit Function
   End If

   On Error Resume Next
                  
   Q1 = "SELECT c.*, Com_Nombre, Reg_Nombre FROM (Adm_NContrib as c "
   Q1 = Q1 & " LEFT JOIN Adm_Comuna ON c.Id_Comuna = Adm_Comuna.Id_Comuna)"
   Q1 = Q1 & " LEFT JOIN Adm_Region ON c.Id_Region = Adm_Region.Id_Region"
   Q1 = Q1 & " WHERE c.NC_Rut='" & lsEmp.ItemData(i) & "'"

   Set Rs = OpenRsDao(lHrDb, Q1)
   
   If Rs Is Nothing Then
      Exit Function
   End If
   
   If Rs.EOF = False Then
      gEmprContab.Rut = Left(lsEmp.ItemData(i), Len(lsEmp.ItemData(i)) - 2)
      gEmprContab.Razon = Trim(vFldDao(Rs("NC_Nombre")) & " " & vFldDao(Rs("NC_Paterno")) & " " & vFldDao(Rs("NC_Materno")))
      
      If Len(gEmprContab.Razon) < 2 Then
         gEmprContab.Razon = Trim(vFldDao(Rs("NC_NomCorto")))
      End If

      gEmprContab.Direccion = vFldDao(Rs("NC_Calle"), True) & " #" & vFldDao(Rs("NC_Nro"))
      
      If Trim(vFldDao(Rs("NC_Depto"))) <> "" Then
         gEmprContab.Direccion = gEmprContab.Direccion & " dpto. " & vFld(Rs("NC_Depto"), True)
      End If
      
      IdContrib = vFldDao(Rs("Id_Contrib"))
      gEmprContab.Telefono = vFldDao(Rs("NC_Fono"))
      gEmprContab.Fax = vFldDao(Rs("NC_Fax"))
      gEmprContab.Provinc = vFldDao(Rs("NC_Ciudad"))
      gEmprContab.Comuna = vFldDao(Rs("Com_Nombre"))
      gEmprContab.Region = vFldDao(Rs("Reg_Nombre"))
      gEmprContab.email = vFldDao(Rs("NC_Correo"))
      gEmprContab.Web = ""
      gEmprContab.Giro = vFldDao(Rs("NC_Giro"))
   
      Import = True
   
      gRc.Rc = vbOK
   
   End If
   
   Call CloseRs(Rs)


   Q1 = "SELECT Adm_Rep_Legal.* FROM Adm_Rep_Legal INNER JOIN Adm_Rep_Contrib ON Adm_Rep_Legal.Id_Rep = Adm_Rep_Contrib.Id_Rep"
   Q1 = Q1 & " WHERE Adm_Rep_Contrib.Id_Contrib=" & IdContrib & " AND Adm_Rep_Legal.Rep_Estado <> 0"
   Q1 = Q1 & " ORDER BY Adm_Rep_Legal.Id_Rep "

   Set Rs = OpenRsDao(lHrDb, Q1)
   
   If Rs Is Nothing Then
      Exit Function
   End If
   
   If Rs.EOF = False Then

      gEmprContab.RutRep = vFldDao(Rs("Rep_Rut"))
      gEmprContab.NomRep = Trim(vFldDao(Rs("Rep_Nombre")) & " " & vFldDao(Rs("Rep_Paterno")) & " " & vFldDao(Rs("Rep_Materno")))

      Import = True
   
      gRc.Rc = vbOK

   End If
   
   Call CloseRs(Rs)
   
End Function

Private Sub Bt_Import_Click()

   If Ls_Emp.ListIndex < 0 Then
      Exit Sub
   End If
   
   If Import() Then
      Unload Me
   End If
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   gRc.Rc = vbCancel

   Set lsEmp = New ClsCombo
   Call lsEmp.SetControl(Ls_Emp)
   
   Call FillEmpHR
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CloseDb(lHrDb)
End Sub

#If LPREMU <> 0 Then

' 26 sep 2017: ahora probamos con DAO 3.6
' Funciona con MS ADO 2.8 Library
'Public Function FillEmpHR() As Boolean
'   Dim Conn As ADODB.Connection
'   Dim Q1 As String
'   Dim Rs As ADODB.Recordset
'   Dim Rc As Long
'   Dim i As Integer
'
'   Rc = 0
'
'   On Error Resume Next
'   Set Conn = New ADODB.Connection
'   lDbPath = gHRPath & "\PAR\BD_HR_admin.mdb"
'   If ExistFile(lDbPath) = False Then
'      MsgBox1 "No se encuentra la base de HR en" & vbCrLf & lDbPath, vbExclamation
'      Exit Function
'   End If
'
'   lHrConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & lDbPath & ";Jet OLEDB:Database Password=" & "20" & "080" & "3hr" & ";"
'
'   Call Conn.Open(lHrConnStr)
'   If Err Then
'      MsgBox "Error H" & Hex(Err) & ", " & Error & NL & lDbPath, vbExclamation
'      FillEmpHR = False
'      Exit Function
'   End If
'
'   Q1 = "SELECT NC_NomCorto, NC_Rut "
'   Q1 = Q1 & " FROM Adm_NContrib"
'   Q1 = Q1 & " ORDER BY NC_NomCorto"
'
'   Set Rs = Conn.Execute(Q1)
'
'   If Rs Is Nothing Then
'      FillEmpHR = False
'      Conn.Close
'      Set Conn = Nothing
'      Exit Function
'   End If
'
'   i = 0
'
'   Do While Not Rs.EOF
'      Call lsEmp.AddItem(vFldADO(Rs("NC_NomCorto")), vFldADO(Rs("NC_Rut")))
'
'      i = i + 1
'
'      Rs.MoveNext
'   Loop
'
'   If i > 0 Then
'      Rc = i
'   End If
'
'   Rs.Close
'   Set Rs = Nothing
'   Conn.Close
'   Set Conn = Nothing
'
'End Function

'Public Function Import() As Boolean
'   Dim i As Integer, Q1 As String
'   Dim Rc As Long
'   Dim IdContrib As Long
'   Dim Conn As ADODB.Connection
'   Dim Rs As ADODB.Recordset
'
'   i = lsEmp.ListIndex
'   Import = False
'
'   If i < 0 Then
'      Exit Function
'   End If
'
'   On Error Resume Next
'   Set Conn = New ADODB.Connection
'
'   If ExistFile(lDbPath) = False Then
'      MsgBox1 "No se encuentra la base de HR en" & vbCrLf & lDbPath, vbExclamation
'      Exit Function
'   End If
'
'   Call Conn.Open(lHrConnStr)
'   If Err Then
'      MsgBox "Error H" & Hex(Err) & ", " & Error & NL & lDbPath, vbExclamation
'      Import = False
'      Exit Function
'   End If
'
'   Q1 = "SELECT Adm_NContrib.*, Com_Nombre, Reg_Nombre FROM (Adm_NContrib "
'   Q1 = Q1 & " INNER JOIN Adm_Comuna ON Adm_NContrib.Id_Comuna = Adm_Comuna.Id_Comuna)"
'   Q1 = Q1 & " INNER JOIN Adm_Region ON Adm_NContrib.Id_Region = Adm_Region.Id_Region"
'   Q1 = Q1 & " WHERE NC_Rut='" & lsEmp.ItemData(i) & "'"
'
'   Set Rs = Conn.Execute(Q1)
'
'   If Rs Is Nothing Then
'      Import = False
'      Conn.Close
'      Set Conn = Nothing
'      Exit Function
'   End If
'
'   If Rs.EOF = False Then
'      gEmprContab.Rut = Left(lsEmp.ItemData(i), Len(lsEmp.ItemData(i)) - 2)
'      gEmprContab.Razon = Trim(vFldADO(Rs("NC_Nombre")) & " " & vFldADO(Rs("NC_Paterno")) & " " & vFldADO(Rs("NC_Materno")))
'
'      gEmprContab.Direccion = vFldADO(Rs("NC_Calle"), True) & " #" & vFldADO(Rs("NC_Nro"))
'
'      If Trim(vFldADO(Rs("NC_Depto"))) <> "" Then
'         gEmprContab.Direccion = gEmprContab.Direccion & " dpto. " & vFldADO(Rs("NC_Depto"), True)
'      End If
'
'      IdContrib = vFldADO(Rs("Id_Contrib"))
'      gEmprContab.Telefono = vFldADO(Rs("NC_Fono"))
'      gEmprContab.Fax = vFldADO(Rs("NC_Fax"))
'      gEmprContab.Ciudad = vFldADO(Rs("NC_Ciudad"))
'      gEmprContab.Comuna = vFldADO(Rs("Com_Nombre"))
'      gEmprContab.Region = vFldADO(Rs("Reg_Nombre"))
'      gEmprContab.email = vFldADO((Rs("NC_Correo")))
'      gEmprContab.Web = ""
'      gEmprContab.Giro = vFld(Rs("NC_Giro"))
'
'      Import = True
'
'      gRc.Rc = vbOK
'
'   End If
'
'   Rs.Close
'   Set Rs = Nothing
'
'
'   Q1 = "SELECT Adm_Rep_Legal.* FROM Adm_Rep_Legal INNER JOIN Adm_Rep_Contrib ON Adm_Rep_Legal.Id_Rep = Adm_Rep_Contrib.Id_Rep"
'   Q1 = Q1 & " WHERE Adm_Rep_Contrib.Id_Contrib=" & IdContrib & " AND Adm_Rep_Legal.Rep_Estado <> 0"
'   Q1 = Q1 & " ORDER BY Adm_Rep_Legal.Id_Rep "
'
'   Set Rs = Conn.Execute(Q1)
'
'   If Rs Is Nothing Then
'      Import = False
'      Conn.Close
'      Set Conn = Nothing
'      Exit Function
'   End If
'
'   If Rs.EOF = False Then
'
'      gEmprContab.RutRep = vFld(Rs("Rep_Rut"))
'      gEmprContab.NomRep = Trim(vFldADO(Rs("Rep_Nombre")) & " " & vFldADO(Rs("Rep_Paterno")) & " " & vFldADO(Rs("Rep_Materno")))
'
'      Import = True
'
'      gRc.Rc = vbOK
'
'   End If
'
'   Rs.Close
'   Set Rs = Nothing
'
'   Conn.Close
'   Set Conn = Nothing
'
'End Function

#End If

Private Sub tx_HR_Change()

End Sub
