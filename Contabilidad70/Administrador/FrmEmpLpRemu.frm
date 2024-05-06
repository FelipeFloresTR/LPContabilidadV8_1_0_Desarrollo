VERSION 5.00
Begin VB.Form FrmEmpLpRemu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas en LpRemu"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "FrmEmpLpRemu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox Ls_EmpCap 
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Bt_Import_todo 
      Caption         =   "Capturar todo"
      Height          =   315
      Left            =   6960
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Capturar"
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox Ls_Emp 
      Height          =   5325
      Left            =   1320
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "FrmEmpLpRemu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lConnStr As String
Private lsEmp As ClsCombo
Private lHrConnStr As String
Private lDbPath As String
Private lsEmpCap As ClsCombo

Public lRc As Integer

Public Function FSelect() As Integer

If lRc = vbCancel Then

  
  FSelect = lRc
Else
 'Me.Show vbModal
  
  FSelect = lRc
End If
 

End Function
Public Function FillEmpHR() As Boolean
   Dim i As Integer
   
   FillEmpHR = False
   
Dim PathDbLpRemu As String
    Dim PathDbLpContab As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   Dim Q1 As String
   Dim RsDao As dao.Recordset
   Dim ConnStr As String
    Dim FNBaseAccess2 As String
   Dim DbAccess2 As Database
   Dim Q2 As String
   Dim RsDao2 As dao.Recordset
   Dim ConnStr2 As String
   Dim bErrMsg As String
   
   Dim Rs1 As Recordset
   Dim Rs2 As Recordset
   
    Dim Rc As Long

  If gDbType = SQL_ACCESS Then
  

   'veamos si existe archivo LPRemu.mdb en el path de la aplicación
   PathDbLpRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   
   If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPRemu.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpRemu & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPRemu.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPRemu.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPRemu.mdb", FNBaseAccess) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpRemu = FNBaseAccess
            
            If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPRemu.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
    
   Set DbAccess = OpenDatabase(PathDbLpRemu, False, False, ConnStr)

   If Err <> 0 Or DbAccess Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpRemu.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
    Else
     FillEmpHR = True
     lRc = vbOK
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(DbAccess, Q1)
   
  If RsDao Is Nothing Then
      Exit Function
  
   End If
   
   i = 0
   
   Do While Not RsDao.EOF
      
      Q1 = "SELECT * FROM Empresas where rut = '" & vFldDao(RsDao("Rut")) & "'"
      Set RsDao2 = OpenRsDao(DbMain, Q1)
      
      If RsDao2.EOF Then
      
      Call lsEmp.AddItem(FmtEmprLs(vFmtRut(vFldDao(RsDao("Rut")) & DV_Rut(vFldDao(RsDao("Rut")))), vFldDao(RsDao("RazonSoc"))), vFldDao(RsDao("Rut")))
                   
      End If

      
      i = i + 1
      Call CloseRs(RsDao2)

      RsDao.MoveNext
   Loop
   
   Call CloseRs(RsDao)
   
   Call CloseDb(DbAccess)
   
   Else 'sql server
   
   
  If OpenMsSqlRemu() = True Then
   
   FillEmpHR = True
   lRc = vbOK
   
    'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(lDbRemu, Q1)
   
    If RsDao Is Nothing Then
       Exit Function
    End If
    
    i = 0
    
    Do While Not RsDao.EOF
    
     Q1 = "SELECT * FROM Empresas where rut = '" & vFldDao(RsDao("Rut")) & "'"
      Set Rs2 = OpenRs(DbMain, Q1)
      
      If Rs2.EOF Then
      
      Call lsEmp.AddItem(FmtEmprLs(vFmtRut(vFldDao(RsDao("Rut")) & DV_Rut(vFldDao(RsDao("Rut")))), vFldDao(RsDao("RazonSoc"))), vFldDao(RsDao("Rut")))
                   
      End If
      
       i = i + 1
       
       RsDao.MoveNext
    Loop
       Call CloseRs(RsDao)
       Call CloseDb(lDbRemu)
        
    Else
        MsgBox1 "Problemas al abrir la base de datos de Remuneraciones.", vbExclamation
        'Call Bt_Cancel_Click
        'Unload FrmEmpLpRemu
        'Exit Functio
      lRc = vbCancel
      
        
    End If
   
   End If
   
   gFrmMain.MousePointer = vbDefault
   
      'FillEmpHR = True
      
   'Call CloseRs(RsDao)
   
End Function

Public Function Import() As Boolean
   Dim i As Integer, Q1 As String
   Dim Rc As Long
   Dim IdContrib As Long
   Dim Rs As dao.Recordset
   Dim PathDbLpContab As String
   Dim PathDbLpRemu As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   
   Dim RsDao As dao.Recordset
   Dim ConnStr As String
   Dim FNBaseAccess2 As String
   Dim DbAccess2 As Database
   Dim Q2 As String
   Dim RsDao2 As dao.Recordset
   Dim ConnStr2 As String
   Dim bErrMsg As String
   Dim Rs1 As Recordset
   Dim Rs2 As Recordset
   
   Dim contador As Integer
   
   contador = 0
   i = 0
   
    On Error Resume Next

    Do While lsEmp.ListCount
       
    If i = lsEmp.ListCount Then
'        Dim x As Integer
'        x = 0
'        For x = 0 To lsEmpCap.ListCount


        'lsEmp.RemoveItem (lsEmpCap.list(x))

         
       
'        Next x
        
       'Call lsEmpCap.SetControl(Ls_Emp)
        
        lsEmpCap.Clear
        lsEmp.Clear
        Call FillEmpHR
        
        MsgBox1 "Proceso finalizado, se capturaron " & contador & " empresas desde LpRemu", vbInformation
   
     Exit Function
    End If
    
    If lsEmp.Selected(i) = True Then
    
        Dim vRut As String
        vRut = lsEmp.ItemData(i)
        
    If gDbType = SQL_ACCESS Then
  
   'veamos si existe archivo LPRemu.mdb en el path de la aplicación
   PathDbLpRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   
   If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPRemu.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpRemu & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPRemu.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPRemu.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPRemu.mdb", FNBaseAccess) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpRemu = FNBaseAccess
            
            If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPRemu.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
    
   Set DbAccess = OpenDatabase(PathDbLpRemu, False, False, ConnStr)

   If Err <> 0 Or DbAccess Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpRemu.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas where rut ='" & vRut & "'"
 
   Set RsDao = OpenRsDao(DbAccess, Q1)
   
   Do While RsDao.EOF = False
   
    'veamos si existe archivo LPContab.mdb en el path de la aplicación
   PathDbLpContab = gDbPath & "\LPContab.mdb"
   If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPContab.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpContab & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPContab.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess2) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpContab = FNBaseAccess2
            
            If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPContab.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr2 = ";PWD=" & PASSW_LEXCONT & ";"
    
   Set DbAccess2 = OpenDatabase(PathDbLpContab, False, False, ConnStr2)

   If Err <> 0 Or DbAccess2 Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpContab.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
   Set RsDao2 = OpenRsDao(DbAccess2, Q1)
   
   If RsDao2.EOF = True Then
       Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
       Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
       Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
       Q1 = Q1 & ", 0, ' ' )"
              
       Call ExecSQL(DbMain, Q1, False)
        'lsEmp.RemoveItem (i)
       
       'Call lsEmpCap.AddItem(FmtEmprLs(vFmtRut(vFldDao(RsDao("Rut")) & DV_Rut(vFldDao(RsDao("Rut")))), vFldDao(RsDao("RazonSoc"))), vFldDao(RsDao("Rut")))
       
       Call lsEmpCap.AddItem(i)
       
       contador = contador + 1
         Else
         Call lsEmpCap.AddItem(FmtEmprLs(vFmtRut(vFldDao(RsDao("Rut")) & DV_Rut(vFldDao(RsDao("Rut")))), vFldDao(RsDao("RazonSoc"))), vFldDao(RsDao("Rut")))
       
    End If
    
    RsDao.MoveNext
   Call CloseRs(RsDao2)
   Call CloseDb(DbAccess2)
   Loop
   
   Call CloseRs(RsDao)
   
   Call CloseDb(DbAccess)
   
    'contador = contador + 1
   
   Else 'sql server
   
   
  If OpenMsSqlRemu() = True Then
   
    'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(lDbRemu, Q1)
   
   Do While RsDao.EOF = False
   
        'leemos la lista de empresas
        Q1 = ""
        Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
        Set Rs2 = OpenRs(DbMain, Q1)
        
        If Rs2.EOF = True Then
            Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
            Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
            Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
            Q1 = Q1 & ", 0, ' ' )"
                   
            Call ExecSQL(DbMain, Q1, False)
              
               lsEmp.RemoveItem (i)
               'Call lsEmpCap.AddItem(i)
              contador = contador + 1
              
           Else
           Call lsEmpCap.AddItem(FmtEmprLs(vFmtRut(vFldDao(RsDao("Rut")) & DV_Rut(vFldDao(RsDao("Rut")))), vFldDao(RsDao("RazonSoc"))), vFldDao(RsDao("Rut")))
       
         End If
        
         RsDao.MoveNext
        Call CloseRs(Rs2)
        'Call CloseDb(DbMain)
        Loop
    
        Call CloseRs(RsDao)
        
        Call CloseDb(lDbRemu)
        
        Else
        MsgBox1 "Problemas al abrir la base de datos de Remuneraciones.", vbExclamation
    
        Exit Function
        
        End If
   
   End If
   
   gFrmMain.MousePointer = vbDefault
   
    End If
  
    
    i = i + 1
    Loop
    
    
   
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

Private Sub Bt_Import_todo_Click()
 Call ImpListEmpLpRemuFromAccess
End Sub

Public Sub Form_Load()

   gRc.Rc = vbCancel
   
   Set lsEmp = New ClsCombo
   Call lsEmp.SetControl(Ls_Emp)
   
   Set lsEmpCap = New ClsCombo
   Call lsEmpCap.SetControl(Ls_EmpCap)
 
   
   'Call ImpListEmpLpRemuFromAccess
  If FillEmpHR = False Then
  
  ' lRc = vbCancel
   
     
  End If
   
  
End Sub


Private Sub Ls_Emp_DblClick()
   Call Bt_Import_Click
End Sub

'ImpListEmpLpRemuFromAccess
'2850275
Public Function ImpListEmpLpRemuFromAccess() As Boolean
   Dim PathDbLpRemu As String
    Dim PathDbLpContab As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   Dim Q1 As String
   Dim RsDao As dao.Recordset
   Dim ConnStr As String
    Dim FNBaseAccess2 As String
   Dim DbAccess2 As Database
   Dim Q2 As String
   Dim RsDao2 As dao.Recordset
   Dim ConnStr2 As String
   Dim bErrMsg As String
   
   Dim Rs1 As Recordset
   Dim Rs2 As Recordset
   
    Dim Rc As Long

  If gDbType = SQL_ACCESS Then
  

   'veamos si existe archivo LPRemu.mdb en el path de la aplicación
   PathDbLpRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   
   If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPRemu.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpRemu & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPRemu.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPRemu.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPRemu.mdb", FNBaseAccess) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpRemu = FNBaseAccess
            
            If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPRemu.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
    
   Set DbAccess = OpenDatabase(PathDbLpRemu, False, False, ConnStr)

   If Err <> 0 Or DbAccess Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpRemu.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(DbAccess, Q1)
   
   Dim contador As Integer
   
   contador = 0
   
   Do While RsDao.EOF = False
   
    'veamos si existe archivo LPContab.mdb en el path de la aplicación
   PathDbLpContab = gDbPath & "\LPContab.mdb"
   If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPContab.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpContab & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPContab.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess2) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpContab = FNBaseAccess2
            
            If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPContab.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr2 = ";PWD=" & PASSW_LEXCONT & ";"
    
   Set DbAccess2 = OpenDatabase(PathDbLpContab, False, False, ConnStr2)

   If Err <> 0 Or DbAccess2 Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpContab.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
   Set RsDao2 = OpenRsDao(DbAccess2, Q1)
   
   If RsDao2.EOF = True Then
       Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
       Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
       Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
       Q1 = Q1 & ", 0, ' ' )"
              
       Call ExecSQL(DbMain, Q1, False)
         
         contador = contador + 1
    End If
    
    RsDao.MoveNext
   Call CloseRs(RsDao2)
   Call CloseDb(DbAccess2)
   Loop
   
   Call CloseRs(RsDao)
   
   Call CloseDb(DbAccess)
   
   
   Else 'sql server
   
   
  If OpenMsSqlRemu() = True Then
   
    'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(lDbRemu, Q1)
   
   Do While RsDao.EOF = False
   
        'leemos la lista de empresas
        Q1 = ""
        Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
        Set Rs2 = OpenRs(DbMain, Q1)
        
        If Rs2.EOF = True Then
            Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
            Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
            Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
            Q1 = Q1 & ", 0, ' ' )"
                   
            Call ExecSQL(DbMain, Q1, False)
              
              contador = contador + 1
         End If
    
         RsDao.MoveNext
        Call CloseRs(Rs2)
        'Call CloseDb(DbMain)
        Loop
   
        Call CloseRs(RsDao)
        
        Call CloseDb(lDbRemu)
        
        Else
        MsgBox1 "Problemas al abrir la base de datos de Remuneraciones.", vbExclamation
    
        Exit Function
        
        End If
   
   End If
   
   gFrmMain.MousePointer = vbDefault
   
   lsEmp.Clear
     
   MsgBox1 "Proceso finalizado, se capturaron " & contador & " empresas desde LpRemu", vbInformation
   
End Function

