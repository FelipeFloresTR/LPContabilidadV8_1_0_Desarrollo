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

Dim lRc As Integer

Dim lHrDb As Database
Public Function FSelect() As Integer

  Me.Show vbModal
  
  FSelect = lRc

End Function
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
      gEmprHR.EmpConta.Rut = Left(lsEmp.ItemData(i), Len(lsEmp.ItemData(i)) - 2)
      gEmprHR.EmpConta.NombreCorto = Trim(vFldDao(Rs("NC_NomCorto")))
      gEmprHR.EmpConta.RazonSocial = Trim(vFldDao(Rs("NC_Nombre")) & " " & vFldDao(Rs("NC_Paterno")) & " " & vFldDao(Rs("NC_Materno")))
      
      gEmprHR.EmpConta.Direccion = vFldDao(Rs("NC_Calle"), True) & " #" & vFldDao(Rs("NC_Nro"))
      
      If Trim(vFldDao(Rs("NC_Depto"))) <> "" Then
         gEmprHR.EmpConta.Direccion = gEmprHR.EmpConta.Direccion & " dpto. " & vFld(Rs("NC_Depto"), True)
      End If
      
      IdContrib = vFldDao(Rs("Id_Contrib"))
      gEmprHR.EmpConta.Telefono = vFldDao(Rs("NC_Fono"))
      gEmprHR.EmpConta.Fax = vFldDao(Rs("NC_Fax"))
      gEmprHR.EmpConta.Ciudad = vFldDao(Rs("NC_Ciudad"))
      gEmprHR.EmpConta.Comuna = vFldDao(Rs("Com_Nombre"))
      gEmprHR.Region = vFldDao(Rs("Reg_Nombre"))
      gEmprHR.EmpConta.email = vFldDao(Rs("NC_Correo"))
'      gEmprHR.Web = ""
      gEmprHR.EmpConta.Giro = vFldDao(Rs("NC_Giro"))
      gEmprHR.EmpConta.Villa = vFldDao(Rs("NC_Villa"))
      gEmprHR.EmpConta.CodArea = vFldDao(Rs("NC_CodArea"))
      gEmprHR.EmpConta.Celular = vFldDao(Rs("NC_Celular"))
   
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

      gEmprHR.EmpConta.RutRepLegal1 = vFldDao(Rs("Rep_Rut"))
      gEmprHR.EmpConta.RepLegal1 = Trim(vFldDao(Rs("Rep_Nombre")) & " " & vFldDao(Rs("Rep_Paterno")) & " " & vFldDao(Rs("Rep_Materno")))

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
      lRc = vbOK
      Unload Me
   End If
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
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


Private Sub Ls_Emp_DblClick()
   Call Bt_Import_Click
End Sub
