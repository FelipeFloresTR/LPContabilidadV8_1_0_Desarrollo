VERSION 5.00
Begin VB.Form FrmIntTools 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_AjustesExtraCont 
      Caption         =   "Ajustes ExtraCont"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reemplaza NewLine en Docs"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   2535
   End
End
Attribute VB_Name = "FrmIntTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_AjustesExtraCont_Click()
   Dim Frm As FrmAjustesExtraLibCaja
   
   Set Frm = New FrmAjustesExtraLibCaja
   Frm.FEdit
   
   Set Frm = Nothing
End Sub

Private Sub Command1_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Txt As String
   Dim i As Integer
   
   Q1 = "SELECT IdMov, Glosa FROM MovComprobante WHERE " & GenLike(DbMain, "½¼", "Glosa")
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 0

   Do While Not Rs.EOF
      i = i + 1

      Txt = Rs("Glosa")
      Txt = ReplaceStr(Txt, "½¼", "")
      Q1 = "UPDATE MovComprobante SET Glosa = '" & Txt & "' WHERE IdMov = " & vFld(Rs("IdMov"))
      Call ExecSQL(DbMain, Q1)
      
      '3376884
      Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmIntTools.Command1_Click", Q1, 1, "WHERE IdMov = " & vFld(Rs("IdMov")), 1, 2)
      'fin 3376884
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   MsgBox1 ("Listo! " & i)
   
End Sub
