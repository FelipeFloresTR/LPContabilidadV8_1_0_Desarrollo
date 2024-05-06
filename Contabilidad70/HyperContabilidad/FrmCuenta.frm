VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuenta Contable"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "FrmCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Hlp14Ter 
      Height          =   4815
      Left            =   8460
      TabIndex        =   40
      Top             =   3180
      Width           =   3105
      Begin MSFlexGridLib.MSFlexGrid Grid14Ter 
         Height          =   4635
         Left            =   60
         TabIndex        =   41
         ToolTipText     =   "Doble-click para seleccionar código"
         Top             =   120
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   8176
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         GridLinesFixed  =   1
         BorderStyle     =   0
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   180
      Picture         =   "FrmCuenta.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   600
      TabIndex        =   33
      Top             =   360
      Width           =   600
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   1020
      TabIndex        =   28
      Top             =   300
      Width           =   8595
      Begin VB.TextBox Tx_CuentaPadre 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "000-000-000-000-000"
         Top             =   180
         Width           =   5835
      End
      Begin VB.Label Lb_TitCuentaPadre 
         AutoSize        =   -1  'True
         Caption         =   "Crear Cuenta bajo:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nueva cuenta"
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Index           =   0
      Left            =   1020
      TabIndex        =   20
      Top             =   1140
      Width           =   8595
      Begin VB.ComboBox Cb_CuentaSII 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2760
         Width           =   5535
      End
      Begin VB.ComboBox Cb_Partida 
         Height          =   315
         Left            =   2100
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   5535
      End
      Begin VB.CommandButton Bt_ConfigIFRS 
         Caption         =   "Configurar Códigos IFRS..."
         Height          =   315
         Left            =   4980
         TabIndex        =   5
         Top             =   1740
         Width           =   2655
      End
      Begin VB.TextBox Tx_CodF29 
         Height          =   315
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   16
         Top             =   3300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Tx_CodF22 
         Height          =   315
         Left            =   6180
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Tx_CodFECU 
         Height          =   315
         Left            =   6180
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   240
         Picture         =   "FrmCuenta.frx":05AF
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   26
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox Tx_Codigo 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "000-000-000-000-000"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox Tx_Descripcion 
         Height          =   315
         Left            =   2100
         MaxLength       =   100
         TabIndex        =   1
         Top             =   900
         Width           =   5535
      End
      Begin VB.TextBox Tx_NombreCorto 
         Height          =   315
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta SII:"
         Height          =   195
         Index           =   9
         Left            =   1020
         TabIndex        =   38
         Top             =   2820
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Partida:"
         Height          =   195
         Index           =   8
         Left            =   1020
         TabIndex        =   36
         Top             =   2340
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Form 29:"
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Form 22:"
         Height          =   195
         Index           =   5
         Left            =   4980
         TabIndex        =   31
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código FECU:"
         Height          =   195
         Index           =   4
         Left            =   5040
         TabIndex        =   27
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   25
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   2
         Left            =   1020
         TabIndex        =   23
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   22
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Corto:"
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   21
         Top             =   1380
         Width           =   1020
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   9960
      TabIndex        =   15
      Top             =   720
      Width           =   1275
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   9960
      TabIndex        =   14
      Top             =   360
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Index           =   1
      Left            =   1020
      TabIndex        =   19
      Top             =   4980
      Width           =   8595
      Begin VB.Frame Fr_14Ter 
         Caption         =   "14 Ter"
         Height          =   915
         Left            =   5100
         TabIndex        =   34
         Top             =   1860
         Width           =   2535
         Begin VB.CommandButton Bt_Ayuda14Ter 
            Caption         =   "?"
            Height          =   315
            Left            =   2040
            TabIndex        =   39
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox Tx_CodF22_14Ter 
            Height          =   315
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Lb_CodF22 
            AutoSize        =   -1  'True
            Caption         =   "Cód.Form 22:"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   420
            Width           =   945
         End
      End
      Begin VB.ListBox Ls_Atrib 
         Height          =   2310
         Left            =   300
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   4215
      End
      Begin VB.Frame Fr_CPropio 
         Caption         =   "Capital Propio"
         Height          =   1095
         Index           =   2
         Left            =   5100
         TabIndex        =   17
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton Op_CPropio 
            Caption         =   "Pasivo Exigible"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   8
            Top             =   300
            Width           =   2115
         End
         Begin VB.OptionButton Op_CPropio 
            Caption         =   "Pasivo No Exigible"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   9
            Top             =   660
            Width           =   2115
         End
      End
      Begin VB.Frame Fr_CPropio 
         Caption         =   "Capital Propio"
         Height          =   1395
         Index           =   1
         Left            =   5100
         TabIndex        =   18
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton Op_CPropio 
            Caption         =   "Valor INTO"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   660
            Width           =   2115
         End
         Begin VB.OptionButton Op_CPropio 
            Caption         =   "Activo Normal"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   2115
         End
         Begin VB.OptionButton Op_CPropio 
            Caption         =   "Complemen. de Activo"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   12
            Top             =   1020
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "FrmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CONCEPTO = 0
Const C_CODIGO = 1

Dim lRc As Integer
Dim lOper As Integer
Dim lCuenta As Cuenta_t

Dim lInLoad As Boolean

Const HEIGHT_SMALL = 5415

Dim lCodHermano As String
Dim lEstado As Integer
Dim lHijos As Integer
Dim lcbCuentaSII As ClsCombo

Dim lCurAtrib(MAX_ATRIB) As Boolean


Private Sub Bt_Ayuda14Ter_Click()

   Fr_Hlp14Ter.visible = Not Fr_Hlp14Ter.visible

End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_ConfigIFRS_Click()
   Dim Frm As FrmConfigCodIFRS

   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmConfigCodIFRS
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_OK_Click()
      
   If Not Valida() Then
      Exit Sub
   End If
   
   Call SaveAll
   
   lRc = vbOK
   Unload Me
End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, j As Integer, Rc As Long
   Dim CapPropio As Integer
   Dim FldArray(6) As AdvTbAddNew_t
   
   On Error Resume Next
   
   If lOper = O_NEW Then
   
      '*** 17-MAY-2005 PAM - Ahora verifica primero si ya existe
      Q1 = "SELECT * FROM Cuentas WHERE Codigo='" & VFmtCodigoCta(Tx_Codigo) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then ' ya existía pero estaba huerfanita
         lCuenta.id = vFld(Rs("idCuenta"))
         Call CloseRs(Rs)
         
         Q1 = "UPDATE Cuentas SET idPadre=" & lCuenta.IdPadre
         Q1 = Q1 & ", Nivel=" & lCuenta.Nivel
         Q1 = Q1 & ", Clasificacion=" & lCuenta.Tipo
         Q1 = Q1 & " WHERE idCuenta=" & lCuenta.id
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Rc = ExecSQL(DbMain, Q1)
      
      Else
         Call CloseRs(Rs)
         
'         Set Rs = DbMain.OpenRecordset("Cuentas", dbOpenTable)
'         Rs.AddNew
'
'         lCuenta.id = Rs("idCuenta")
'         Rs("idPadre") = lCuenta.IdPadre
'         Rs("Codigo") = VFmtCodigoCta(Tx_Codigo)
'         Rs("Nivel") = lCuenta.Nivel
'         'Rs("Clasificacion") = lCuenta.Tipo
'         Rs("Clasificacion") = CbItemData(Cb_Tipo)     'FCA: 17/05/2006 lCuenta.Tipo viene con valor 0 cuando es de primer nivel
'         Rs("TipoPartida") = CbItemData(Cb_Partida)
'         Rs.Update
'         Rs.Close
         

         FldArray(0).FldName = "Codigo"
         FldArray(0).FldValue = VFmtCodigoCta(Tx_Codigo)
         FldArray(0).FldIsNum = True
         
         FldArray(1).FldName = "idPadre"
         FldArray(1).FldValue = lCuenta.IdPadre
         FldArray(1).FldIsNum = True
               
         FldArray(2).FldName = "Nivel"
         FldArray(2).FldValue = lCuenta.Nivel
         FldArray(2).FldIsNum = True
               
         FldArray(3).FldName = "Clasificacion"
         FldArray(3).FldValue = lCuenta.Tipo
         FldArray(3).FldIsNum = True
               
         FldArray(4).FldName = "TipoPartida"
         FldArray(4).FldValue = CbItemData(Cb_Partida)
         FldArray(4).FldIsNum = True
               
         FldArray(5).FldName = "IdEmpresa"
         FldArray(5).FldValue = gEmpresa.id
         FldArray(5).FldIsNum = True
                     
         FldArray(6).FldName = "Ano"
         FldArray(6).FldValue = gEmpresa.Ano
         FldArray(6).FldIsNum = True
         
         lCuenta.id = AdvTbAddNewMult(DbMain, "Cuentas", "IdCuenta", FldArray)
         
'         Q1 = "UPDATE Cuentas SET "
'         Q1 = Q1 & "  idPadre = " & lCuenta.IdPadre
'         Q1 = Q1 & ", Nivel = " & lCuenta.Nivel
'         Q1 = Q1 & ", Clasificacion = " & lCuenta.Tipo
'         Q1 = Q1 & ", TipoPartida = " & CbItemData(Cb_Partida)
'         Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
'         Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'
'         Q1 = Q1 & " WHERE IdCuenta = " & lCuenta.id
'         Call ExecSQL(DbMain, Q1)
         
         If ERR Then
            MsgErr "Error al crear cuenta"
            Exit Sub
         End If
         
      End If
      
   End If
   
   
   
   
   'Ahora actualizamos los otros campos
      
   Q1 = "UPDATE Cuentas SET "
   Q1 = Q1 & " CodFECU='" & Trim(Tx_CodFECU) & "'"
   Q1 = Q1 & ",CodF22=" & Val(Tx_CodF22)
   Q1 = Q1 & ",CodF29=" & Val(Tx_CodF29)
   Q1 = Q1 & ",Nombre='" & ParaSQL(Tx_NombreCorto) & "'"
   Q1 = Q1 & ",Descripcion='" & ParaSQL(Tx_Descripcion) & "'"
   Q1 = Q1 & ",Estado= 1"
   Q1 = Q1 & ",TipoPartida= " & CbItemData(Cb_Partida)
   Q1 = Q1 & ",CodCtaPlanSII= '" & ParaSQL(lcbCuentaSII.Matrix(2)) & "'"  'CodigoSII
   
   'atributos
   For i = 0 To Ls_Atrib.ListCount - 2
      Q1 = Q1 & ", Atrib" & Ls_Atrib.ItemData(i) & "=" & Abs(Ls_Atrib.Selected(i))
   Next i
   
   Q1 = Q1 & ", Percepcion" & "=" & Abs(Ls_Atrib.Selected(Ls_Atrib.ListCount - 1))
   
   'capital propio y 14 TER

   For i = 0 To Ls_Atrib.ListCount - 1
      If Ls_Atrib.ItemData(i) = ATRIB_CAPITALPROPIO Then
         If Ls_Atrib.Selected(i) = True Then
            For j = 1 To MAX_CAPPROPIO
               If Op_CPropio(j) = True Then
                  Q1 = Q1 & ", TipoCapPropio=" & j
                  Exit For
               End If
            Next j
         Else
            Q1 = Q1 & ", TipoCapPropio = 0"
         
         End If
         
      ElseIf Ls_Atrib.ItemData(i) = ATRIB_14TER Then
         If Ls_Atrib.Selected(i) = True Then
            Q1 = Q1 & ", CodF22_14Ter=" & vFmt(Tx_CodF22_14Ter)
            lCuenta.CodF22_14Ter = Val(Tx_CodF22_14Ter)
         Else
            Q1 = Q1 & ", CodF22_14Ter=0"
            lCuenta.CodF22_14Ter = 0
         End If
      End If
   Next i
      
   Q1 = Q1 & " WHERE idCuenta=" & lCuenta.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'cambiamos tipo de cuenta de los hijos, si hay cambio
   If lCuenta.Nivel = NIVEL_1 Then
      If lCuenta.Tipo <> ItemData(Cb_Tipo) Then
         Q1 = "UPDATE Cuentas SET Clasificacion = " & ItemData(Cb_Tipo)
         Q1 = Q1 & " WHERE left(Codigo," & gNiveles.Largo(NIVEL_1) & ") = '" & Left(lCuenta.Codigo, gNiveles.Largo(NIVEL_1)) & "'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
   End If
   
   'actualizamos variable global
   lCuenta.Codigo = Tx_Codigo
   lCuenta.CodFECU = Trim(Tx_CodFECU)
   lCuenta.CodF22 = Val(Tx_CodF22)
   lCuenta.CodF29 = Val(Tx_CodF29)
   lCuenta.Descripcion = Tx_Descripcion
   lCuenta.Tipo = ItemData(Cb_Tipo)
   lCuenta.Nombre = Tx_NombreCorto
   
   'limpiamos los atributos del padre si corresponde
   If lCuenta.Nivel = gLastNivel Then   'el padre podría tener atributos definidos
      
      Q1 = "UPDATE Cuentas SET "
      Q1 = Q1 & "Atrib1 = 0"
      For i = 2 To MAX_ATRIB
         Q1 = Q1 & ", Atrib" & i & "=0"
      Next i
      
      Q1 = Q1 & " WHERE idCuenta=" & lCuenta.IdPadre
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
   End If

End Sub
Private Function Valida() As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Atrib As Integer

   Valida = False
   
   If Cb_Tipo.ListIndex < 0 Then
      Call MsgBox1("Falta seleccionar la clasificación de la cuenta.", vbExclamation + vbOKOnly)
      Exit Function
   End If
   
   If lCuenta.Nivel = 1 Then
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Nivel = 1 AND Clasificacion = " & ItemData(Cb_Tipo) & " AND IdCuenta <> " & lCuenta.id
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         If ItemData(Cb_Tipo) <> CLASCTA_RESULTADO Then
            MsgBox1 "Ya existe una cuenta de primer nivel con esta clasificación.", vbExclamation
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
      Call CloseRs(Rs)
   End If
   
   If lOper = O_EDIT And lCuenta.Tipo <> ItemData(Cb_Tipo) Then
      If MsgBox1("¿Está seguro que desea cambiar la clasificación de esta cuenta?" & vbNewLine & vbNewLine & "Atención: este cambio también se realizará en todas las cuentas que se encuentran bajo esta cuenta.", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
      
   If Trim(Tx_Descripcion) = "" Then
      Call MsgBox1("Falta ingresar la descripción de la cuenta.", vbExclamation + vbOKOnly)
      Exit Function
   End If
   
    Q1 = "SELECT Franq14ASemiIntegrado FROM EMPRESA "
    Q1 = Q1 & " WHERE Rut = '" & gEmpresa.Rut & "' AND Ano = " & gEmpresa.Ano
    Set Rs = OpenRs(DbMain, Q1)
    If Not Rs.EOF Then
       If CbItemData(Me.Cb_CuentaSII) = 0 Then
        If vFld(Rs("Franq14ASemiIntegrado")) <> 0 Then
           If MsgBox1("ID Cuenta SII no tiene clasificación… desea Continuar?.", vbQuestion + vbYesNo) = vbNo Then
             Exit Function
           End If
        End If
       End If
    End If
    Call CloseRs(Rs)
   
   If gEmpresa.Ano = 2017 Then
      If Trim(Tx_CodF22) <> "" Then
         If InStr("," & LSTCODF22_2017 & ",", "," & Trim(Tx_CodF22) & ",") > 0 Then     'es inválido
            MsgBox1 "Código Form 22 inválido." & vbCrLf & vbCrLf & "Desde el año 2017 los siguientes códigos no son válidos:" & vbCrLf & LSTCODF22_2017, vbExclamation
            Exit Function
         End If
      End If
   End If
      
   Tx_NombreCorto = Trim(Tx_NombreCorto)
   If Tx_NombreCorto <> "" Then
      'el nombre corto no es obligatorio, pero si lo pone, debe ser único
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Nombre = '" & ParaSQL(Tx_NombreCorto) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If vFld(Rs(0)) <> lCuenta.id Then
            Call MsgBox1("Este nombre corto ya existe para otra cuenta.", vbExclamation + vbOKOnly)
            Call CloseRs(Rs)
            Exit Function
         End If
      End If
      
      Call CloseRs(Rs)
   End If
   
   Valida = True
   
End Function
Public Function FView(ByVal IdCuenta As Long) As Integer
   lOper = O_VIEW
   Me.Show vbModal
   
   FView = lRc
End Function
Friend Function FEdit(Cuenta As Cuenta_t) As Integer
   lOper = O_EDIT
   lCuenta = Cuenta
   Me.Show vbModal
   
   Cuenta = lCuenta
   FEdit = lRc
End Function

Friend Function FNew(Cuenta As Cuenta_t) As Integer
   lOper = O_NEW
   lCuenta = Cuenta
   lCuenta.id = 0
   Me.Show vbModal
   
   Cuenta = lCuenta
   FNew = lRc
   
End Function
Private Sub Cb_Tipo_Click()

   If CbItemData(Cb_Tipo) = CLASCTA_RESULTADO And lCuenta.Nivel = gLastNivel Then
      Cb_Partida.Enabled = True
      Tx_CodF22 = ""
      Tx_CodF22.Enabled = False
   Else
      Cb_Partida.Enabled = False
      If Cb_Partida.ListCount > 0 Then
         Cb_Partida.ListIndex = 0
      End If
   End If

   If lCuenta.Nivel = gLastNivel Then
      Cb_CuentaSII.Enabled = True
   Else
      Cb_CuentaSII.Enabled = False
      If Cb_CuentaSII.ListCount > 0 Then
         Cb_CuentaSII.ListIndex = 0
      End If
   End If

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   lRc = vbCancel
   
   lInLoad = True
   
   'Lleno las Combos
             
   'If lCuenta.Nivel = NIVEL_1 And lOper = O_NEW Then
   If lCuenta.Nivel = NIVEL_1 Then
      For i = 1 To MAX_CLASCTA
         Cb_Tipo.AddItem gClasCta(i)
         Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = i
         
         If lOper <> O_NEW And lCuenta.Tipo = i Then
            Cb_Tipo.ListIndex = Cb_Tipo.NewIndex
         End If
               
      Next i
      
      If lOper = O_NEW Then
         Cb_Tipo.ListIndex = -1
      End If
      
   Else
   
      Cb_Tipo.AddItem gClasCta(lCuenta.Tipo)
      Cb_Tipo.ItemData(Cb_Tipo.NewIndex) = lCuenta.Tipo
      Cb_Tipo.ListIndex = 0
      
   End If
      
   'lista de atributos
   For i = 1 To MAX_ATRIB
      Ls_Atrib.AddItem gAtribCuentas(i).Nombre
      Ls_Atrib.ItemData(Ls_Atrib.NewIndex) = i
      
           
          '2961932
           If Not gEmpresa.Ano <= 2019 And Not gEmpresa.Franq14Ter Then
           ' tema 4 2738156
               If Not gEmpresa.ProPymeGeneral And Not gEmpresa.ProPymeTransp Then
    
                If Ls_Atrib.ItemData(Ls_Atrib.NewIndex) = ATRIB_14TER Then
                    Ls_Atrib.RemoveItem (Ls_Atrib.NewIndex)
                    
                End If
                
             ' fin tema 4 2738156
           End If
           '2961932
      End If
   Next i
   
   
    
   
   Ls_Atrib.ListIndex = 0
   
   'Tipo Partida
   Call CbAddItem(Cb_Partida, "", 0)
   For i = 1 To MAX_TIPOPARTIDA
      If (gTipoPartida(i).AnoDesde = 0 Or gEmpresa.Ano >= gTipoPartida(i).AnoDesde) And (gTipoPartida(i).AnoHasta = 0 Or gEmpresa.Ano <= gTipoPartida(i).AnoHasta) And (gTipoPartida(i).SoloArt14A = 0 Or Not gEmpresa.R14ASemiIntegrado) Then
         Call CbAddItem(Cb_Partida, gTipoPartida(i).Partida, i)
      End If
   Next i

   Cb_Partida.ListIndex = 0
   
   'Cuenta SII
   Set lcbCuentaSII = New ClsCombo
   Call lcbCuentaSII.SetControl(Cb_CuentaSII)

   Call lcbCuentaSII.AddItem("")
   If lCuenta.Nivel = gLastNivel Then
      Q1 = "SELECT FmtCodigoSII + ' - ' + DescripSII, IdPlanCuentasSII, CodigoSII FROM PlanCuentasSII "
      Q1 = Q1 & " WHERE Clasificacion = " & lCuenta.Tipo
      Q1 = Q1 & " AND (AnoDesde IS NULL OR AnoDesde = 0 or AnoDesde <= " & gEmpresa.Ano & ")"
      Q1 = Q1 & " ORDER BY CodigoSII "
      Call lcbCuentaSII.FillCombo(DbMain, Q1, "0")
   End If
   lcbCuentaSII.ListIndex = 0
   
   Fr_CPropio(CLASCTA_ACTIVO).visible = False
   Fr_CPropio(CLASCTA_PASIVO).visible = False
   Fr_14Ter.visible = False
   Fr_Hlp14Ter.visible = False
   Call FillHlp14Ter
   
         
   If lOper = O_NEW Then
      'FORMO EL NUEVO CODIGO
      Tx_Codigo = NewCodigoCta(lCuenta, lCodHermano)
      Caption = "Nueva " & Caption
      Lb_TitCuentaPadre = "Crear Cuenta bajo: "
      Tx_CuentaPadre = GetPathCuenta(lCuenta.IdPadre)
      LoadNew
   Else
      Caption = "Modifica " & Caption
      Lb_TitCuentaPadre = "Modificar Cuenta: "
      Tx_CuentaPadre = GetPathCuenta(lCuenta.id)
      LoadEdit
   End If
      
   If Tx_Codigo = "" Then
      Bt_OK.Enabled = False
      Call SetTxRO(Tx_Descripcion, True)
      'Call SettxRO(Tx_CodF22, True)
      Call SetTxRO(Tx_CodF29, True)
      Call SetTxRO(Tx_CodFECU, True)
      Call SetTxRO(Tx_NombreCorto, True)
      Ls_Atrib.Enabled = False
   Else
      Tx_NombreCorto.MaxLength = Len(VFmtCodigoCta(Tx_Codigo)) - 1
      If Tx_NombreCorto.MaxLength > 10 Then
         Tx_NombreCorto.MaxLength = 10
      End If
   End If
        
   'por ahora no dejamos editar códigos formularios
   'Call SettxRO(Tx_CodF22, True)
   Call SetTxRO(Tx_CodF29, True)
   
   If gEmpresa.Ano >= 2020 Then
      Fr_14Ter.Caption = "Art. 14D N°3 y 8 LIR"
      Lb_CodF22 = "Cód. Ajustes"
   End If
   
 
   
   Call SetupPriv
   
   lInLoad = False
   
End Sub
Private Sub LoadEdit()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim RemoveIdx As Integer
            
   'VEO SI ES PADRE DE ALGUNA CUENTA
   Q1 = "SELECT Count(idPadre) as Hijos FROM Cuentas "
   Q1 = Q1 & " WHERE idPadre=" & lCuenta.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   lHijos = vFld(Rs("Hijos"))
   Call CloseRs(Rs)
  
   If lCuenta.Nivel <> gLastNivel Then
      Me.Height = HEIGHT_SMALL
      Bt_ConfigIFRS.visible = False
   End If
   
   Q1 = "SELECT Codigo, CodFECU, CodF22, CodF29, idPadre, Nombre, Descripcion, Nivel, Estado, Clasificacion, TipoPartida, TipoCapPropio, CodF22_14Ter, CodCtaPlanSII "
   For i = 1 To MAX_ATRIB
      Q1 = Q1 & ",Atrib" & i
   Next i
   Q1 = Q1 & ",Percepcion "
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE idCuenta=" & lCuenta.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   RemoveIdx = -1
   
   If Rs.EOF = False Then
      Tx_Codigo = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Tx_CodFECU = vFld(Rs("CodFECU"))
      
      If vFld(Rs("CodF22")) > 0 Then
         Tx_CodF22 = vFld(Rs("CodF22"))
      End If
      If vFld(Rs("CodF29")) > 0 Then
         Tx_CodF29 = vFld(Rs("CodF29"))
      End If
      
      Tx_Descripcion = vFld(Rs("Descripcion"), True)
      Tx_NombreCorto = vFld(Rs("Nombre"), True)
      lEstado = vFld(Rs("Estado"))
      If ItemData(Cb_Tipo) <> CLASCTA_ACTIVO And ItemData(Cb_Tipo) <> CLASCTA_PASIVO Then
         Tx_CodF22 = ""
         Call SetRO(Tx_CodF22, True)
      End If
      
      If lCuenta.Nivel = gLastNivel Then
      
         For i = 0 To Ls_Atrib.ListCount - 1
            Ls_Atrib.Selected(i) = (vFld(Rs("Atrib" & Ls_Atrib.ItemData(i))) <> 0)
            lCurAtrib(i) = Ls_Atrib.Selected(i)
            Ls_Atrib.Selected(Ls_Atrib.ListCount - 1) = (vFld(Rs("Percepcion")) <> 0)
            lCurAtrib(Ls_Atrib.ListCount - 1) = Ls_Atrib.Selected(Ls_Atrib.ListCount - 1)
            
            If Ls_Atrib.ItemData(i) = ATRIB_CAPITALPROPIO Then
               
               If ItemData(Cb_Tipo) <> CLASCTA_ACTIVO And ItemData(Cb_Tipo) <> CLASCTA_PASIVO Then
                  RemoveIdx = i
               
               ElseIf vFld(Rs("TipoCapPropio")) > 0 And vFld(Rs("TipoCapPropio")) <= MAX_CAPPROPIO Then
                  Op_CPropio(vFld(Rs("TipoCapPropio"))) = True
               End If
            
            ElseIf Ls_Atrib.ItemData(i) = ATRIB_14TER Then
            
               If Ls_Atrib.Selected(i) And vFld(Rs("CodF22_14Ter")) > 0 Then
                  Tx_CodF22_14Ter = HomologaCod14D(vFld(Rs("CodF22_14Ter")))
               End If
            End If
            
         Next i
         
         Ls_Atrib.ListIndex = 0
         
         If RemoveIdx >= 0 Then
            Ls_Atrib.RemoveItem (RemoveIdx)   'eliminamos capital propio si corresponde
         End If
         
         Call CbSelItem(Cb_Partida, vFld(Rs("TipoPArtida")))
         lcbCuentaSII.ListIndex = lcbCuentaSII.FindItem(vFld(Rs("CodCtaPlanSII")), 2)

      End If
      
    
   End If
   Call CloseRs(Rs)
   
End Sub

Private Sub LoadNew()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim RemoveIdx As Integer
   Dim Wh As String
   
   RemoveIdx = -1
   
   If lCuenta.Nivel <> gLastNivel Then
      Me.Height = HEIGHT_SMALL
   
   Else
   
      'si no es cuenta de activo o pasivo, eliminamos atributo Capital Propio
      
      If ItemData(Cb_Tipo) <> CLASCTA_ACTIVO And ItemData(Cb_Tipo) <> CLASCTA_PASIVO Then
         Call SetRO(Tx_CodF22, True)
         Tx_CodF22 = ""
         
         For i = 0 To Ls_Atrib.ListCount - 1
            
            If Ls_Atrib.ItemData(i) = ATRIB_CAPITALPROPIO Then
               RemoveIdx = i
            End If
            
         Next i
      
         If RemoveIdx >= 0 Then
            Ls_Atrib.RemoveItem (RemoveIdx)   'eliminamos capital propio si corresponde
         End If
           
      End If
   
      'Si no tiene hermanos hereda los atributos de su padre, siempre que este
      'no sea el nivel 1, ya que no tendría qué heredar.
      'Si tiene hermanos, hereda atributos del hermano
      If lCodHermano = "" And lCuenta.NivelFather <> NIVEL_1 Then
         Wh = " WHERE idCuenta=" & lCuenta.IdPadre
      ElseIf lCodHermano <> "" Then
         Wh = " WHERE Codigo = '" & lCodHermano & "'"
      End If
         
      If Wh <> "" Then
      
         Wh = Wh & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
         Q1 = "SELECT CodFECU, CodF22, CodF29, Descripcion, "
         For i = 1 To MAX_ATRIB
            Q1 = Q1 & "Atrib" & i & ","
         Next i
         Q1 = Left(Q1, Len(Q1) - 1)
         Q1 = Q1 & " FROM Cuentas "
         Q1 = Q1 & Wh
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            Tx_CodFECU = vFld(Rs("CodFECU"))
            If vFld(Rs("CodF22")) > 0 Then
               Tx_CodF22 = vFld(Rs("CodF22"))
            End If
            If vFld(Rs("CodF29")) > 0 Then
               Tx_CodF29 = vFld(Rs("CodF29"))
            End If
            Tx_Descripcion = vFld(Rs("Descripcion"))
         
            For i = 0 To Ls_Atrib.ListCount - 1
               Ls_Atrib.Selected(i) = (vFld(Rs("Atrib" & Ls_Atrib.ItemData(i))) <> 0)
            Next i
            
         End If
         
         Call CloseRs(Rs)
         
      End If
     
   End If
   
End Sub


Private Sub Grid14Ter_DblClick()

   If Grid14Ter.Row > 0 Then
      Tx_CodF22_14Ter = Val(Grid14Ter.TextMatrix(Grid14Ter.Row, C_CODIGO))
   End If
   
End Sub

Private Sub Ls_Atrib_Click()
   Dim i As Integer
   Static InLsAtrib As Boolean
   
   If InLsAtrib Then
      Exit Sub
   End If
   
   If Ls_Atrib.ListIndex < 0 Then
      Exit Sub
   End If
   
   InLsAtrib = True
   
   If Ls_Atrib.ItemData(Ls_Atrib.ListIndex) = ATRIB_CAPITALPROPIO Then
      
      If Ls_Atrib.Selected(Ls_Atrib.ListIndex) = True Then
         If Not lInLoad Then
            For i = 1 To MAX_CAPPROPIO
               Op_CPropio(i) = False
            Next i
         End If
         
         If ItemData(Cb_Tipo) = CLASCTA_ACTIVO Then
            Fr_CPropio(CLASCTA_ACTIVO).visible = True
            Fr_CPropio(CLASCTA_PASIVO).visible = False
            If Not lInLoad Then
               Op_CPropio(CAPPROPIO_ACTIVO_NORMAL) = True
            End If
            
         ElseIf ItemData(Cb_Tipo) = CLASCTA_PASIVO Then
            Fr_CPropio(CLASCTA_ACTIVO).visible = False
            Fr_CPropio(CLASCTA_PASIVO).visible = True
            If Not lInLoad Then
               Op_CPropio(CAPPROPIO_PASIVO_EXIGIBLE) = True
            End If
         Else
            Fr_CPropio(CLASCTA_ACTIVO).visible = False
            Fr_CPropio(CLASCTA_PASIVO).visible = False
         End If
         
      Else
         Fr_CPropio(CLASCTA_ACTIVO).visible = False
         Fr_CPropio(CLASCTA_PASIVO).visible = False
            
      End If
   
   ElseIf Ls_Atrib.ItemData(Ls_Atrib.ListIndex) = ATRIB_14TER Then
      If Ls_Atrib.Selected(Ls_Atrib.ListIndex) = True Then
         Fr_14Ter.visible = True
      Else
         Fr_14Ter.visible = False
         Tx_CodF22_14Ter = ""
      End If
      
      DoEvents
      
   ElseIf Not lInLoad And (Ls_Atrib.ItemData(Ls_Atrib.ListIndex) = ATRIB_RUT Or Ls_Atrib.ItemData(Ls_Atrib.ListIndex) = ATRIB_CCOSTO Or Ls_Atrib.ItemData(Ls_Atrib.ListIndex) = ATRIB_AREANEG) Then
      
      If CuentaTieneMovs(lCuenta.id) Then
      
         If MsgBox1("Usted esta modificando los atributos de una cuenta contable que posee información. Si modifica los atributos podría generar diferencias entre los distintos reportes que entrega el sistema." & vbCrLf & vbCrLf & "¿Está seguro de efectuar este cambio?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Ls_Atrib.Selected(Ls_Atrib.ListIndex) = lCurAtrib(Ls_Atrib.ListIndex)
         End If
      End If
      
   End If
      
   InLsAtrib = False
   

End Sub


Private Sub Tx_CodF22_14Ter_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_CodF22_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_CodF22_LostFocus()

   Tx_CodF22 = Trim(Tx_CodF22)

   If gEmpresa.Ano = 2017 Then
      If Trim(Tx_CodF22) <> "" Then
         If InStr("," & LSTCODF22_2017 & ",", "," & Trim(Tx_CodF22) & ",") > 0 Then     'es inválido
            MsgBox1 "Código Form 22 inválido." & vbCrLf & vbCrLf & "Desde el año 2017 los siguientes códigos no son válidos:" & vbCrLf & LSTCODF22_2017, vbExclamation
         End If
      End If
   End If

End Sub

Private Sub Tx_NombreCorto_KeyPress(KeyAscii As Integer)
   Call KeyUserId(KeyAscii)
   Call KeyUpper(KeyAscii)
   
End Sub

Private Sub Tx_NombreCorto_LostFocus()
   Tx_NombreCorto = UCase(Tx_NombreCorto)
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_CTAS) Then
      Call EnableForm(Me, False)
   End If
   
End Function

Private Function CuentaTieneMovs(ByVal IdCuenta As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Count(*) FROM MovComprobante WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   CuentaTieneMovs = False
   
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then
         CuentaTieneMovs = True
      End If
   End If
   
   Call CloseRs(Rs)
   
End Function
Private Function FillHlp14Ter()

   Call FGrSetup(Grid14Ter)
   
   Grid14Ter.rows = 15
   
   Grid14Ter.Height = 3615
   Fr_Hlp14Ter.Height = 3795
   Fr_Hlp14Ter.Top = 4200

   
   Grid14Ter.ColWidth(C_CONCEPTO) = 2500
   Grid14Ter.ColWidth(C_CODIGO) = 500
   
   Grid14Ter.ColAlignment(C_CODIGO) = flexAlignRightCenter
   
   Grid14Ter.TextMatrix(0, C_CONCEPTO) = "Detalle"
   Grid14Ter.TextMatrix(0, C_CODIGO) = "Cód."
   
   Grid14Ter.TextMatrix(1, C_CONCEPTO) = "Ingresos Perc. Giro"
   Grid14Ter.TextMatrix(1, C_CODIGO) = HomologaCod14D(628)
   
   Grid14Ter.TextMatrix(2, C_CONCEPTO) = "Participaciones"
   Grid14Ter.TextMatrix(2, C_CODIGO) = HomologaCod14D(629)

   Grid14Ter.TextMatrix(3, C_CONCEPTO) = "Otros Ingr. Percib."
   Grid14Ter.TextMatrix(3, C_CODIGO) = HomologaCod14D(651)

   If gEmpresa.Ano < 2020 Then
      Grid14Ter.TextMatrix(4, C_CONCEPTO) = "Costo Dir. Bienes o Serv."
   Else
      Grid14Ter.TextMatrix(4, C_CONCEPTO) = "Costo Dir. Bienes"
   End If
   Grid14Ter.TextMatrix(4, C_CODIGO) = HomologaCod14D(630)

   Grid14Ter.TextMatrix(5, C_CONCEPTO) = "Remuneraciones"
   Grid14Ter.TextMatrix(5, C_CODIGO) = HomologaCod14D(631)

   Grid14Ter.TextMatrix(6, C_CONCEPTO) = "Activo Fijo"
   Grid14Ter.TextMatrix(6, C_CODIGO) = HomologaCod14D(632)

   Grid14Ter.TextMatrix(7, C_CONCEPTO) = "Intereses Pagados"
   Grid14Ter.TextMatrix(7, C_CODIGO) = HomologaCod14D(633)

   Grid14Ter.TextMatrix(8, C_CONCEPTO) = "Otros Gastos"
   Grid14Ter.TextMatrix(8, C_CODIGO) = HomologaCod14D(635)

      
   If gEmpresa.R14ASemiIntegrado Then
   
      Grid14Ter.rows = 19
      
      Grid14Ter.Height = 4635
      Fr_Hlp14Ter.Height = 4785
      Fr_Hlp14Ter.Top = 3215

      Grid14Ter.TextMatrix(9, C_CONCEPTO) = "Rentas Fuente Extranjera"
      Grid14Ter.TextMatrix(9, C_CODIGO) = HomologaCod14D(851)

      Grid14Ter.TextMatrix(10, C_CONCEPTO) = "Gastos por Donaciones"
      Grid14Ter.TextMatrix(10, C_CODIGO) = HomologaCod14D(966)
   
      Grid14Ter.TextMatrix(11, C_CONCEPTO) = "Otros Gastos Financieros"
      Grid14Ter.TextMatrix(11, C_CODIGO) = HomologaCod14D(967)

      Grid14Ter.TextMatrix(12, C_CONCEPTO) = "Gastos Invest. Des. cert. Corfo"
      Grid14Ter.TextMatrix(12, C_CODIGO) = HomologaCod14D(852)
      
      Grid14Ter.TextMatrix(13, C_CONCEPTO) = "Gastos Invest. Des. No cert. Corfo"
      Grid14Ter.TextMatrix(13, C_CODIGO) = HomologaCod14D(897)

      Grid14Ter.TextMatrix(14, C_CONCEPTO) = "Arriendos"
      Grid14Ter.TextMatrix(14, C_CODIGO) = HomologaCod14D(1140)
   
      Grid14Ter.TextMatrix(15, C_CONCEPTO) = "Gastos exigencias medioambient."
      Grid14Ter.TextMatrix(15, C_CODIGO) = HomologaCod14D(1141)
   
      Grid14Ter.TextMatrix(16, C_CONCEPTO) = "Gastos Indeminz. Comp. Clientes"
      Grid14Ter.TextMatrix(16, C_CODIGO) = HomologaCod14D(1142)
      
      Grid14Ter.TextMatrix(17, C_CONCEPTO) = "Gastos Producir Renta Fte. Ext."
      Grid14Ter.TextMatrix(17, C_CODIGO) = HomologaCod14D(1669)
   
      Grid14Ter.TextMatrix(18, C_CONCEPTO) = "Gastos Imp. Renta e Imp. Diferido"
      Grid14Ter.TextMatrix(18, C_CODIGO) = HomologaCod14D(1670)
   
      
   Else
      Grid14Ter.TextMatrix(9, C_CONCEPTO) = "INR Propios"
      Grid14Ter.TextMatrix(9, C_CODIGO) = HomologaCod14D(640)

      Grid14Ter.TextMatrix(10, C_CONCEPTO) = "Aumentos de Capital"
      Grid14Ter.TextMatrix(10, C_CODIGO) = HomologaCod14D(893)
   
      Grid14Ter.TextMatrix(11, C_CONCEPTO) = "Disminuciones Capital"
      Grid14Ter.TextMatrix(11, C_CODIGO) = HomologaCod14D(894)

      Grid14Ter.TextMatrix(12, C_CONCEPTO) = "Gastos Rech. no Afec. 21"
      Grid14Ter.TextMatrix(12, C_CODIGO) = HomologaCod14D(990)

      Grid14Ter.TextMatrix(13, C_CONCEPTO) = "Arriendos"
      Grid14Ter.TextMatrix(13, C_CODIGO) = HomologaCod14D(1140)
      
      Grid14Ter.TextMatrix(14, C_CONCEPTO) = "Part. Inc. 1° no afec. IU tasa 40%"
      Grid14Ter.TextMatrix(14, C_CODIGO) = HomologaCod14D(1144)
      
   End If
   

End Function
