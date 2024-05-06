VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LP Contabilidad"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7380
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   8325
      Left            =   0
      Picture         =   "FrmAbout.frx":000C
      ScaleHeight     =   8265
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      Begin VB.TextBox Tx_Ubicacion 
         BackColor       =   &H00AD7900&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   260
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Ubicaci"
         Top             =   6300
         Width           =   6615
      End
      Begin VB.Label Lb_AccessSQL 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2220
         TabIndex        =   13
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Lb_Version 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tel/Fax: (56 2) 2483 8600"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   10
         Top             =   5160
         Width           =   7275
      End
      Begin VB.Label la_Link 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "email  soporte.chile@thomsonreuters.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   3480
         MouseIcon       =   "FrmAbout.frx":6374
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Tag             =   "mailto:soporte@legalpublishing.cl?Subject=Legal Publishing%20Contabilidad"
         Top             =   5520
         Width           =   3255
      End
      Begin VB.Label la_Link 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.legalpublishing.cl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   540
         MouseIcon       =   "FrmAbout.frx":64C6
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Tag             =   "http://www.legalpublishing.cl"
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   60
         TabIndex        =   7
         Top             =   4560
         Width           =   7275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Desarrollado por:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   6
         Top             =   3840
         Width           =   7275
      End
      Begin VB.Label La_Ver 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 00.00.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   3300
         Width           =   1455
      End
      Begin VB.Label La_Fecha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00 mmm 0000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Thomson Reuters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   4140
         Width           =   7275
      End
      Begin VB.Label La_Nivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1860
         TabIndex        =   2
         Top             =   3300
         Width           =   3675
      End
      Begin VB.Label Lb_Demo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEMO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5760
         TabIndex        =   1
         Top             =   1980
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'prueba 5

'pipe prueba
'pruebas pipe 12-10-22
'Prueba Fer 12-10-22
'prueba simultanea
Private Sub Form_Load()
    Dim Buf As String, Q1 As String
    Dim Dt As Long
    Dim i As Integer
    Dim Rs As Recordset
    

    Me.Icon = FrmMain.Icon
    'Im_Icon.Picture = Me.Icon
    Lb_Demo.visible = gAppCode.Demo

    ' Image1.Picture = Me.Icon
    ' la_Link(0).MouseIcon = FrmMain.Fr_Invivisible.MouseIcon
    ' la_Link(1).MouseIcon = FrmMain.Fr_Invivisible.MouseIcon

    Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
        Buf = "/" & Val(vFld(Rs("Valor")))
    Else
        Buf = ""
    End If
    Call CloseRs(Rs)

    Tx_Ubicacion = "Ubicación: " & W.AppPath
    '   Lb_ubicacion = "Ubicación: " & W.AppPath

    Me.Caption = "Acerca de " & gLexContab
    'La_Title = gLexContab

    '   La_Ver = "Versión " & W.Version & Buf
    La_Ver = "V " & App.Major & "." & App.Minor & "." & App.Revision & Buf
    Lb_Version = "Versión " & App.Major & "." & App.Minor
    Lb_AccessSQL = IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")

    La_Nivel = gAppCode.NivProd
    For i = 0 To UBound(gAppCode.Nivel)
        If gAppCode.Nivel(i).id = gAppCode.NivProd Then
            La_Nivel = gAppCode.Nivel(i).Desc
        End If
    Next i

    '   Select Case gAppCode.NivProd
    '      Case VMANT_2005
    '         La_Nivel = "c/Mant. 2005"
    '
    '      Case 1:
    '         La_Nivel = "Básico"
    '
    '      Case Else:
    '         La_Nivel = "¿" & gAppCode.NivProd & "?"
    '
    '   End Select

    La_Fecha = Format(W.FVersion, "mmm d, yyyy")

End Sub

Private Sub OK_Click()
   Unload Me
End Sub
Private Sub la_Link_Click(Index As Integer)
   Dim Rc As Long
   Dim Buf As String
   
   If la_Link(Index).Tag <> "" Then
      Buf = la_Link(Index).Tag
   Else
      Buf = la_Link(Index)
   End If
   
   Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", SW_SHOWNORMAL)
   
End Sub

Private Sub Tx_Ubicacion_DblClick()
   Dim i As Integer, Rc As Long

   Clipboard.Clear
   i = InStr(Tx_Ubicacion, ":") ' Ubicación:

   If i > 0 Then
      Clipboard.SetText Trim(Mid(Tx_Ubicacion, i + 1))
      Call ShellExecute(Me.hWnd, "open", Trim(Mid(Tx_Ubicacion, i + 1)), "", "", SW_SHOWNORMAL)
   Else
      Clipboard.SetText Trim(Tx_Ubicacion)
   End If

End Sub

