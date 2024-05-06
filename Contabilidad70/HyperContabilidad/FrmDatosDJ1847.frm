VERSION 5.00
Begin VB.Form FrmDatosDJ1847 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información requerida para Exportar DJ 1847"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese los siguientes antecedentes para continuar"
      Height          =   1395
      Left            =   1260
      TabIndex        =   8
      Top             =   360
      Width           =   8895
      Begin VB.ComboBox Cb_AjustesRLI 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox Tx_FolioFinal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Tx_FolioInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3660
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Tx_AnoAjusteIFRS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox Cb_EntSupervisora 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ajustes RLI*"
         Height          =   195
         Index           =   4
         Left            =   6420
         TabIndex        =   13
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio Final*"
         Height          =   195
         Index           =   3
         Left            =   5040
         TabIndex        =   12
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio Inicial*"
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   11
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año Ajuste IFRS"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entidad Supervisora*"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.CommandButton Bt_Exp 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   7500
      TabIndex        =   5
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8880
      TabIndex        =   6
      Top             =   1920
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   360
      Picture         =   "FrmDatosDJ1847.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   585
      TabIndex        =   7
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "(*) dato obligatorio"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   1980
      Width           =   1635
   End
End
Attribute VB_Name = "FrmDatosDJ1847"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer

Public Function FEdit() As Integer
   Me.Show vbModal
   FEdit = lRc
End Function

Private Sub Bt_Close_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Exp_Click()
   If Valida() Then
      Call SaveAll
      lRc = vbOK
      Unload Me
   End If
End Sub

Private Function Valida() As Boolean
   Dim MaxFolio As Double

   Valida = False
   
   If CbItemData(Cb_EntSupervisora) <= 0 Then
      MsgBox1 "Falta seleccionar la Entidad Supervisora.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_AnoAjusteIFRS) = 0 And Tx_AnoAjusteIFRS <> "0000" Then
      MsgBox1 "Año ajuste IFRS inválido. Puede usar '0000' si no tiene la aplicación de IFRS.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_AnoAjusteIFRS) > 2020 Then
      MsgBox1 "Año ajuste IFRS inválido. No puede ser porterior al año 2020.", vbExclamation
      Exit Function
   End If

   If vFmt(Tx_FolioInicial) <= 0 Then
      MsgBox1 "Folio inicial inválido.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_FolioFinal) <= 0 Then
      MsgBox1 "Folio final inválido.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_FolioInicial) > vFmt(Tx_FolioFinal) Then
      MsgBox1 "Folio inicial debe ser menor o igual a folio final.", vbExclamation
      Exit Function
   End If
   
   MaxFolio = 9999999999#
   
   If vFmt(Tx_FolioInicial) > MaxFolio Then
      MsgBox1 "Folio inicial debe ser menor o igual a 9.999.999.999.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_FolioFinal) > MaxFolio Then
      MsgBox1 "Folio final debe ser menor o igual a " & Format(MaxFolio, NUMFMT), vbExclamation
      Exit Function
   End If
   
   If CbItemData(Cb_AjustesRLI) <= 0 Then
      MsgBox1 "Falta seleccionar la Entidad Supervisora.", vbExclamation
      Exit Function
   End If
   
   Valida = True
End Function
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT IdEntSupervisora, AnoAjusteIFRS, FolioInicial, FolioFinal, IdAjustesRLI "    'esta tabla tiene sólo 1 registro por empresa/ano
   Q1 = Q1 & " FROM InfoAnualDJ1847 WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Call CbSelItem(Cb_EntSupervisora, vFld(Rs("IdEntSupervisora")))
      Tx_AnoAjusteIFRS = IIf(vFld(Rs("AnoAjusteIFRS")) > 0, vFld(Rs("AnoAjusteIFRS")), "")
      Tx_FolioInicial = Format(vFld(Rs("FolioInicial")), NUMFMT)
      Tx_FolioFinal = Format(vFld(Rs("FolioFinal")), NUMFMT)
      Call CbSelItem(Cb_AjustesRLI, vFld(Rs("IdAjustesRLI")))
   End If
   
   Call CloseRs(Rs)
   

End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Count(*) FROM InfoAnualDJ1847 WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
  
   If Not Rs.EOF Then
      If vFld(Rs(0)) = 0 Then   'esta tabla tiene sólo 1 registro por empresa/ano
         Q1 = "INSERT INTO InfoAnualDJ1847 (IdEmpresa, Ano) VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & ")"
         Call ExecSQL(DbMain, Q1)
      End If
   End If
   Call CloseRs(Rs)
   
   Q1 = "UPDATE InfoAnualDJ1847 SET "        'esta tabla tiene sólo 1 registro por empresa/ano
   Q1 = Q1 & "  IdEntSupervisora = " & CbItemData(Cb_EntSupervisora)
   Q1 = Q1 & ", AnoAjusteIFRS = " & vFmt(Tx_AnoAjusteIFRS)
   Q1 = Q1 & ", FolioInicial = " & vFmt(Tx_FolioInicial)
   Q1 = Q1 & ", FolioFinal = " & vFmt(Tx_FolioFinal)
   Q1 = Q1 & ", IdAjustesRLI = " & CbItemData(Cb_AjustesRLI)
   Q1 = Q1 & "  WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
     
End Sub

Private Sub Form_Load()
   Dim i As Integer

   If gDJ1847_EntSupervisora(1) = "" Then
      Call InitDJ1847
   End If
      
   For i = 1 To UBound(gDJ1847_EntSupervisora)
      
      If gEmpresa.Ano >= 2020 Then
         If i <> DJ1847_ES_SVS And i <> DJ1847_ES_SBIF Then
            Call CbAddItem(Cb_EntSupervisora, gDJ1847_EntSupervisora(i), i)
         End If
      Else
         If i <> DJ1847_ES_CMF Then
            Call CbAddItem(Cb_EntSupervisora, gDJ1847_EntSupervisora(i), i)
         End If
      End If
      
   Next i
   Cb_EntSupervisora.ListIndex = -1
      
   For i = 1 To UBound(gDJ1847_AjusteRLI)
      Call CbAddItem(Cb_AjustesRLI, gDJ1847_AjusteRLI(i), i)
   Next i
   Cb_AjustesRLI.ListIndex = -1
   
   Call LoadAll
   
End Sub

Private Sub Tx_AnoAjusteIFRS_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_FolioFinal_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_FolioFinal_LostFocus()
   Tx_FolioFinal = Format(vFmt(Tx_FolioFinal), NUMFMT)

End Sub

Private Sub Tx_FolioInicial_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_FolioInicial_LostFocus()
   Tx_FolioInicial = Format(vFmt(Tx_FolioInicial), NUMFMT)
   
End Sub
