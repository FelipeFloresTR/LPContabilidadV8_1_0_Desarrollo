VERSION 5.00
Begin VB.Form FrmConfigInformeFirma 
   Caption         =   "Configuraciones de Firma para Informes"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Notas"
      Height          =   1215
      Left            =   1080
      TabIndex        =   16
      Top             =   2520
      Width           =   7935
      Begin VB.Label Label4 
         Caption         =   "La imagen de la firma debe tener las siguientes dimensiones Ancho 228 y Alto 112 (228x112 px)."
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6315
      End
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imágenes de Firmas"
      Height          =   2250
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton Bt_Del 
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
         Index           =   2
         Left            =   7320
         Picture         =   "FrmConfigInformeFirma.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar documento seleccionado"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
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
         Index           =   1
         Left            =   7320
         Picture         =   "FrmConfigInformeFirma.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Eliminar documento seleccionado"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
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
         Index           =   0
         Left            =   7320
         Picture         =   "FrmConfigInformeFirma.frx":07F8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar documento seleccionado"
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Pb_RepLegal2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   2400
         ScaleHeight     =   135
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Pb_RepLegal1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1440
         ScaleHeight     =   135
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox pb_contador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   480
         ScaleHeight     =   135
         ScaleWidth      =   615
         TabIndex        =   13
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Bt_SearchRepLegal2 
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
         Left            =   6840
         Picture         =   "FrmConfigInformeFirma.frx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Buscar Firma Rep. Legal"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Bt_SearchRepLegal1 
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
         Left            =   6840
         Picture         =   "FrmConfigInformeFirma.frx":0F77
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar Firma Rep. Legal"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_SearchContador 
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
         Left            =   6840
         Picture         =   "FrmConfigInformeFirma.frx":12FA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Buscar Firma Contador"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Tx_RepLegal2 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   5535
      End
      Begin VB.TextBox Tx_RepLegal1 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox Tx_Contador 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rep. Legal 2:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rep. Legal 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contador:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   2
      Left            =   240
      Picture         =   "FrmConfigInformeFirma.frx":167D
      ScaleHeight     =   705
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "FrmConfigInformeFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Bt_Aceptar_Click()

Dim Folder As String
Dim Q1 As String
Dim Rc As Long

If Tx_Contador = "" And Tx_RepLegal1 = "" And Tx_RepLegal2 = "" Then
    If MsgBox1("No hay firmas cargadas, desea continuar.", vbInformation + vbYesNo) = vbYes Then
       
       Exit Sub
    Else
       Unload Me
       Exit Sub
    End If
    
 Else
 


Folder = App.Path & "\Firma"

 ' Si no existe la carpeta, la creamos
 If Dir(Folder, vbDirectory) = "" Then
     MkDir (Folder)
 Else
 
    If pb_contador.Picture <> 0 Then
    pb_contador.ScaleMode = vbPixels
    ' Set scale to pixels.
    pb_contador.AutoRedraw = True    ' If needed
    
    pb_contador.Width = 3475  ' in pixels
    pb_contador.Height = 1830  ' in pixels

    SavePicture pb_contador.Image, Folder & "\" & gEmpresa.Rut & "_firContador_" & gEmpresa.Ano & ".jpg"
     
        If ExisteFirma("Contador") Then
        Q1 = "UPDATE Firmas SET patch = '" & Folder & "\" & gEmpresa.Rut & "_firContador_" & gEmpresa.Ano & ".jpg'"
        Q1 = Q1 & " WHERE Tipo = 'Contador' and idEmpresa= " & gEmpresa.id & " and ano = '" & gEmpresa.Ano & "'"
        Rc = ExecSQL(DbMain, Q1)
        
        Else
        Q1 = "INSERT INTO Firmas (patch,  IdEmpresa,Tipo,ano) VALUES ('" & Folder & "\" & gEmpresa.Rut & "_firContador_" & gEmpresa.Ano & ".jpg" & "'," & gEmpresa.id & ",'Contador','" & gEmpresa.Ano & "')"
        Rc = ExecSQL(DbMain, Q1)
        End If
    
     Bt_Del(0).Enabled = True
    End If
    
    If Pb_RepLegal1.Picture <> 0 Then
    Pb_RepLegal1.ScaleMode = vbPixels
    ' Set scale to pixels.
    Pb_RepLegal1.AutoRedraw = True    ' If needed
    
    Pb_RepLegal1.Width = 3475  ' in pixels
    Pb_RepLegal1.Height = 1830  ' in pixels
     SavePicture Pb_RepLegal1.Image, Folder & "\" & gEmpresa.Rut & "_firRepLegal1_" & gEmpresa.Ano & ".jpg"
    
    
      If ExisteFirma("RepLegal1") Then
      Q1 = "UPDATE Firmas SET patch = '" & Folder & "\" & gEmpresa.Rut & "_firRepLegal1_" & gEmpresa.Ano & ".jpg'"
      Q1 = Q1 & " WHERE Tipo = 'RepLegal1' and idEmpresa= " & gEmpresa.id & " and ano = '" & gEmpresa.Ano & "'"
      Rc = ExecSQL(DbMain, Q1)
      Else
      
      Q1 = "INSERT INTO Firmas (patch,  IdEmpresa,Tipo,ano) VALUES ('" & Folder & "\" & gEmpresa.Rut & "_firRepLegal1_" & gEmpresa.Ano & ".jpg" & "'," & gEmpresa.id & ",'RepLegal1','" & gEmpresa.Ano & "')"
      Rc = ExecSQL(DbMain, Q1)
      End If
      Bt_Del(1).Enabled = True
     End If
     If Pb_RepLegal2.Picture <> 0 Then
        Pb_RepLegal2.ScaleMode = vbPixels
        ' Set scale to pixels.
        Pb_RepLegal2.AutoRedraw = True    ' If needed
        
        Pb_RepLegal2.Width = 3475  ' in pixels
        Pb_RepLegal2.Height = 1830  ' in pixels
         SavePicture Pb_RepLegal2.Image, Folder & "\" & gEmpresa.Rut & "_firRepLegal2_" & gEmpresa.Ano & ".jpg"
        
         If ExisteFirma("RepLegal2") Then
         Q1 = "UPDATE Firmas SET patch = '" & Folder & "\" & gEmpresa.Rut & "_firRepLegal2_" & gEmpresa.Ano & ".jpg'"
         Q1 = Q1 & " WHERE Tipo = 'RepLegal2' and idEmpresa= " & gEmpresa.id & " and ano = '" & gEmpresa.Ano & "'"
         Rc = ExecSQL(DbMain, Q1)
         Else
           
         Q1 = "INSERT INTO Firmas (patch,  IdEmpresa,Tipo,ano) VALUES ('" & Folder & "\" & gEmpresa.Rut & "_firRepLegal2_" & gEmpresa.Ano & ".jpg" & "'," & gEmpresa.id & ",'RepLegal2','" & gEmpresa.Ano & "')"
         Rc = ExecSQL(DbMain, Q1)
              
         End If
     Bt_Del(2).Enabled = True
     End If
        If ERR = cdlCancel Then
        Exit Sub
        ElseIf ERR Then
           MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
           Exit Sub
        End If
        
         MsgBox1 "Firmas cargadas exitosamente.", vbInformation
   Unload Me
      
 End If
            
End If
End Sub

Private Sub Bt_Cancelar_Click()
Unload Me
End Sub

Private Sub Bt_Del_Click(Index As Integer)
Dim Q1 As String
Dim Q2 As String
Dim Rc As Long
Dim Rs As Recordset

  Q1 = "DELETE FROM Firmas "

Select Case Index
    Case 0
                       
            Q2 = "Select patch from Firmas where idempresa = " & gEmpresa.id & " And Tipo= 'Contador' and ano='" & gEmpresa.Ano & "'"

            Set Rs = OpenRs(DbMain, Q2)
            
            If Rs.EOF = False Then
             Kill (vFld(Rs("patch")))
            End If
             Call CloseRs(Rs)
             
             Q1 = Q1 & "where idEmpresa= " & gEmpresa.id & " And Tipo= 'Contador' and ano='" & gEmpresa.Ano & "'"
            Rc = ExecSQL(DbMain, Q1)
             
             MsgBox1 "Firma Contador Eliminada Correctamente.", vbInformation
             
             Tx_Contador = ""
             Bt_Del(0).Enabled = False
             
    Case 1
                      
           Q2 = "Select patch from Firmas where idempresa = " & gEmpresa.id & " And Tipo= 'RepLegal1' and ano='" & gEmpresa.Ano & "'"

            Set Rs = OpenRs(DbMain, Q2)
            
            If Rs.EOF = False Then
             Kill (vFld(Rs("patch")))
            End If
             Call CloseRs(Rs)
             
             Q1 = Q1 & "where idEmpresa= " & gEmpresa.id & " And Tipo= 'RepLegal1' and ano='" & gEmpresa.Ano & "'"
            Rc = ExecSQL(DbMain, Q1)
             MsgBox1 "Firma Rep. Legal 1 Eliminada Correctamente.", vbInformation
             
             Tx_RepLegal1 = ""
             Bt_Del(1).Enabled = False
    Case 2
            
            Q2 = "Select patch from Firmas where idempresa = " & gEmpresa.id & " And Tipo= 'RepLegal2' and ano='" & gEmpresa.Ano & "'"

            Set Rs = OpenRs(DbMain, Q2)
            
            If Rs.EOF = False Then
             Kill (vFld(Rs("patch")))
            End If
             Call CloseRs(Rs)
             
             Q1 = Q1 & "where idEmpresa= " & gEmpresa.id & " And Tipo= 'RepLegal2' and ano='" & gEmpresa.Ano & "'"
            Rc = ExecSQL(DbMain, Q1)
            MsgBox1 "Firma Rep. Legal 2 Eliminada Correctamente.", vbInformation
            
            Tx_RepLegal2 = ""
            Bt_Del(2).Enabled = False
    Case Else
    
End Select
 
   

End Sub

Private Sub Bt_SearchContador_Click()

Set pb_contador = Nothing

 gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = ""
  
      gFrmMain.Cm_ComDlg.Filter = "Image Files(*.JPG;*.JPEG)|*.JPG;*.JPEG"
      gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar imagen Firma Contador"
   
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   
  pb_contador.AutoRedraw = True
   pb_contador.Picture = LoadPicture(FrmMain.Cm_ComDlg.Filename)
   
   If ImagenAncho(pb_contador.Picture) > 228 And ImagenAlto(pb_contador.Picture) > 112 Then
     MsgBox1 "Imagen supera la dimensiones sugeridas 228x112 px", vbExclamation
      Tx_Contador = ""
      pb_contador = Nothing
  Else
   Tx_Contador = FrmMain.Cm_ComDlg.Filename
   
   End If
   
  
     ERR.Clear

End Sub

Private Sub Bt_SearchRepLegal1_Click()
Set Pb_RepLegal1 = Nothing

 gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = ""
  
      gFrmMain.Cm_ComDlg.Filter = "*.jpg"
      gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar imagen Firma Rep. Legal 1"
   
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   
  Pb_RepLegal1.AutoRedraw = True
   Pb_RepLegal1.Picture = LoadPicture(FrmMain.Cm_ComDlg.Filename)
   
    If ImagenAncho(Pb_RepLegal1.Picture) > 228 And ImagenAlto(Pb_RepLegal1.Picture) > 112 Then
      MsgBox1 "Imagen supera la dimensiones sugeridas 228x112 px", vbExclamation
      Tx_RepLegal1 = ""
      Pb_RepLegal1 = Nothing
    Else
      Tx_RepLegal1 = FrmMain.Cm_ComDlg.Filename
      
    End If
   
   
     ERR.Clear

End Sub

Private Sub Bt_SearchRepLegal2_Click()
Set Pb_RepLegal2 = Nothing

 gFrmMain.Cm_ComDlg.CancelError = True
   gFrmMain.Cm_ComDlg.Filename = ""
  
      gFrmMain.Cm_ComDlg.Filter = "*.jpg"
      gFrmMain.Cm_ComDlg.DialogTitle = "Seleccionar imagen Firma Rep. Legal 2"
   
   gFrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoChangeDir
 
   On Error Resume Next
   gFrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   
  Pb_RepLegal2.AutoRedraw = True
   Pb_RepLegal2.Picture = LoadPicture(FrmMain.Cm_ComDlg.Filename)
   
   If ImagenAncho(Pb_RepLegal2.Picture) > 228 And ImagenAlto(Pb_RepLegal2.Picture) > 112 Then
      MsgBox1 "Imagen supera la dimensiones sugeridas 228x112 px", vbExclamation
      Tx_RepLegal2 = ""
      Pb_RepLegal2 = Nothing
    Else
      Tx_RepLegal2 = FrmMain.Cm_ComDlg.Filename
      
    End If
   
     ERR.Clear

End Sub

Private Sub Form_Load()
Dim Q1 As String
Dim Rs As Recordset

If gEmpresa.RutContador = "" Then
Tx_Contador.Enabled = False
Bt_SearchContador.Enabled = False
Bt_Del(0).Enabled = False
End If
If gEmpresa.RutRepLegal1 = "" Or gEmpresa.RutRepLegal1 = "0" Then
Tx_RepLegal1.Enabled = False
Bt_SearchRepLegal1.Enabled = False
Bt_Del(1).Enabled = False
End If
If gEmpresa.RutRepLegal2 = "" Or gEmpresa.RutRepLegal2 = "0" Then
Tx_RepLegal2.Enabled = False
Bt_SearchRepLegal2.Enabled = False
Bt_Del(2).Enabled = False
End If

If Tx_Contador.Enabled = False And Tx_RepLegal1.Enabled = False And Tx_RepLegal2.Enabled = False Then
Bt_Aceptar.Enabled = False
End If

Q1 = ""
Q1 = Q1 & "Select patch,idempresa,tipo from Firmas where idempresa = " & gEmpresa.id & " and ano ='" & gEmpresa.Ano & "'"

Set Rs = OpenRs(DbMain, Q1)

Do While Rs.EOF = False
      If vFld(Rs("Tipo")) = "Contador" Then
         Tx_Contador = vFld(Rs("patch"))
      ElseIf vFld(Rs("Tipo")) = "RepLegal1" Then
        Tx_RepLegal1 = vFld(Rs("patch"))
      ElseIf vFld(Rs("Tipo")) = "RepLegal2" Then
        Tx_RepLegal2 = vFld(Rs("patch"))
      End If
      
    Rs.MoveNext
      
Loop
   Call CloseRs(Rs)
   
   If Tx_Contador = "" Then
   Bt_Del(0).Enabled = False
   End If
   If Tx_RepLegal1 = "" Then
   Bt_Del(1).Enabled = False
   End If
   If Tx_RepLegal2 = "" Then
   Bt_Del(2).Enabled = False
   End If

End Sub


Private Function ExisteFirma(ByVal Tipo As String) As Boolean
Dim Q1 As String
Dim Rs As Recordset

ExisteFirma = False

Q1 = ""
Q1 = "Select * from Firmas where idEmpresa =" & gEmpresa.id
Q1 = Q1 & " And Tipo ='" & Tipo & "'"

Set Rs = OpenRs(DbMain, Q1)

If Not Rs.EOF Then
ExisteFirma = True
End If

Call CloseRs(Rs)
End Function

Public Function ImagenAncho(ByRef p As Picture) As Integer
    ImagenAncho = p.Width / Screen.TwipsPerPixelX * 0.566893424036281
End Function


' Devuelve el alto real de un Picture en píxeles
Public Function ImagenAlto(ByRef p As Picture) As Integer
    ImagenAlto = p.Height / Screen.TwipsPerPixelY * 0.566893424036281
End Function


