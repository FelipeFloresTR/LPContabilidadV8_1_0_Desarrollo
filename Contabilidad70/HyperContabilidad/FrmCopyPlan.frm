VERSION 5.00
Begin VB.Form FrmCopyPlan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copiar plan de cuenta de otra empresa"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "FrmCopyPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   180
      Picture         =   "FrmCopyPlan.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   600
      TabIndex        =   6
      Top             =   420
      Width           =   600
   End
   Begin VB.ListBox Ls_Empresas 
      Height          =   2595
      Left            =   1080
      TabIndex        =   0
      Top             =   660
      Width           =   3795
   End
   Begin VB.CommandButton Bt_Cerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5340
      TabIndex        =   2
      Top             =   420
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Copy 
      Caption         =   "&Copiar "
      Height          =   855
      Left            =   5340
      Picture         =   "FrmCopyPlan.frx":05AF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   420
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT"
      Height          =   315
      Index           =   6
      Left            =   1080
      TabIndex        =   4
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Corto"
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   420
      Width           =   2115
   End
End
Attribute VB_Name = "FrmCopyPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lsEmpresas As ClsCombo
Dim lCopyObj As Integer

Const C_COPYPLAN = 1
Const C_COPYCONFIGREMU = 2
Const C_COPYCONFIGIMPADIC = 3

Const MATRIX_MAXANO = 2
Const MATRIX_IDEMP = 3


Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub CopyPlan()
   Dim Q1 As String
   Dim Rc As Long
   Dim Rs As Recordset
   Dim ExistePlan As Boolean
   Dim ConnStr As String
   Dim nNiv As Integer, nNivCopy As Integer
   Dim FldLst As String, AtribLst As String, Fld As String, Fld2 As String
   Dim IdEmpresaCopy As Long, AnoCopy As Integer
   Dim i As Integer
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   '3360107
   Dim nNiv2 As Integer, Nivel As Integer
   '3360107
   If lsEmpresas.ListIndex < 0 Then
      Exit Sub
   End If
   
   
   IdEmpresaCopy = lsEmpresas.Matrix(MATRIX_IDEMP)
   AnoCopy = lsEmpresas.Matrix(MATRIX_MAXANO)
   
   
   'Chequeamos que los niveles definidos en las dos empresas sean los mismos
   'Para esto linkeamos ParamEmpresa
#If DATACON = 1 Then       'Access
  
    '2868088
     ConnStr = "PWD=" & PASSW_PREFIX & lsEmpresas.ItemData & ";"
     'ConnStr = "PWD=" & PASSW_PREFIX_NEW & lsEmpresas.ItemData & ";"
     
      'Call SetDbSecurity(gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb", PASSW_PREFIX & lsEmpresas.ItemData, gCfgFile, SG_SEGCFG, gEmpresa.ConnStr)
     
     'ConnStr = gEmpresa.ConnStr
    '2868088
   If gEmprSeparadas Then
      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb", "ParamEmpresa", "ParamEmpresaCopy", , , ConnStr)
   End If
#End If

   'vemos los niveles definidos para la empresa
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'NIVELES'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      nNiv = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
   'Ahora vemos los nivels de la empresa de origen
   If gEmprSeparadas Then
      Q1 = "SELECT Valor FROM ParamEmpresaCopy WHERE Tipo = 'NIVELES'"   'aquí no se requiere empresa y año
   Else
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'NIVELES'"
      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresaCopy & " AND Ano = " & AnoCopy
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      nNivCopy = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
   If nNiv <> nNivCopy Then
      If MsgBox1("La cantidad de niveles del plan de la empresa no coinciden con los niveles del plan que desea copiar." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
         Exit Sub
      End If
   End If
   
   'Chequeamos que la empresa destino no tenga plan pre-definido
   Q1 = "SELECT Count(*) as Cant FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExistePlan = True
      End If
   End If
   Call CloseRs(Rs)
   
   If ExistePlan Then
      If MsgBox1("Esta empresa ya tiene un plan de cuentas definido. Si copia otro plan, perderá el actual." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
      Call DeleteSQL(DbMain, "Cuentas", " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
   End If
   
   Me.MousePointer = vbHourglass
   
   'copiamos Niveles de ParamEmpresaCopy a esta
   
   If gEmprSeparadas Then
'      Q1 = "UPDATE ParamEmpresa INNER JOIN ParamEmpresaCopy ON ParamEmpresa.Tipo = ParamEmpresaCopy.Tipo"
'      Q1 = Q1 & " SET ParamEmpresa.Valor = ParamEmpresaCopy.Valor WHERE ParamEmpresaCopy.Tipo IN( 'NIVELES', 'DIGNIV1', 'DIGNIV2', 'DIGNIV3', 'DIGNIV4', 'DIGNIV5')"
'      Q1 = Q1 & " AND ParamEmpresaCopy.IdEmpresa = " & IdEmpresaCopy & " AND ParamEmpresaCopy.Ano = " & AnoCopy
      Tbl = " ParamEmpresa "
      sFrom = " ParamEmpresa INNER JOIN ParamEmpresaCopy ON ParamEmpresa.Tipo = ParamEmpresaCopy.Tipo "
      sSet = " ParamEmpresa.Valor = ParamEmpresaCopy.Valor "
      sWhere = " WHERE ParamEmpresaCopy.Tipo IN( 'NIVELES', 'DIGNIV1', 'DIGNIV2', 'DIGNIV3', 'DIGNIV4', 'DIGNIV5')"
      sWhere = sWhere            ' & " AND ParamEmpresaCopy.IdEmpresa = " & IdEmpresaCopy & " AND ParamEmpresaCopy.Ano = " & AnoCopy
   
   Else
'      Q1 = "UPDATE ParamEmpresa INNER JOIN ParamEmpresa as ParamEmpresaCopy ON ParamEmpresa.Tipo = ParamEmpresaCopy.Tipo"
'      Q1 = Q1 & " AND ParamEmpresa.IdEmpresa = " & gEmpresa.id & " AND ParamEmpresa.Ano = " & gEmpresa.Ano
'      Q1 = Q1 & " AND ParamEmpresaCopy.IdEmpresa = " & IdEmpresaCopy & " AND ParamEmpresaCopy.Ano = " & AnoCopy
      Tbl = " ParamEmpresa "
      sFrom = " ParamEmpresa INNER JOIN ParamEmpresa as ParamEmpresaCopy ON ParamEmpresa.Tipo = ParamEmpresaCopy.Tipo"
      sFrom = sFrom & " AND ParamEmpresa.IdEmpresa = " & gEmpresa.id & " AND ParamEmpresa.Ano = " & gEmpresa.Ano
      sFrom = sFrom & " AND ParamEmpresaCopy.IdEmpresa = " & IdEmpresaCopy & " AND ParamEmpresaCopy.Ano = " & AnoCopy
      sSet = " ParamEmpresa.Valor = ParamEmpresaCopy.Valor "
      sWhere = " WHERE ParamEmpresaCopy.Tipo IN( 'NIVELES', 'DIGNIV1', 'DIGNIV2', 'DIGNIV3', 'DIGNIV4', 'DIGNIV5')"
   
   End If
   
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
         
   For i = 1 To MAX_ATRIB
      AtribLst = AtribLst & ", Atrib" & i
   Next i
   
   'FldLst = "Cuentas.idCuenta, Cuentas.idPadre, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion, Cuentas.CodFECU, Cuentas.Nivel, Cuentas.Estado, Cuentas.Clasificacion, Cuentas.Debe, Cuentas.Haber, Cuentas.MarcaApertura, Cuentas.TipoCapPropio, Cuentas.CodF22 " & AtribLst & ", Cuentas.CodF29, Cuentas.CorrelativoCheque, Cuentas.CodIFRS_EstRes, Cuentas.CodIFRS_EstFin, Cuentas.DebeTrib, Cuentas.HaberTrib, Cuentas.CodIFRS, Cuentas.CodF22_14Ter, Cuentas.TipoPartida, Cuentas.CodCtaPlanSII"
   FldLst = "idCuenta, idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22 " & AtribLst & ", CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII, Percepcion"
   
   If gEmprSeparadas Then

#If DATACON = 1 Then       'Access

      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb", "Cuentas", "CuentasCopy", , , ConnStr)
      
'       '3360107
'       'Chequeamos que plan de cuentas a copiar no este con conflictos
        Dim NCta As String, LCta As String

        Q1 = "SELECT " & FldLst & "," & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
        Q1 = Q1 & " FROM CuentasCopy order by codigo asc"
        Set Rs = OpenRs(DbMain, Q1)
        Do While Rs.EOF = False

           If Len(vFld(Rs("codigo"))) > 6 Then

            NCta = GetCodCta(FmtCodCuenta(vFld(Rs("Codigo"))), Nivel, nNiv, gNiveles.Largo())

                If IIf(IsNumeric(Replace(vFld(Rs("codigo")), "-", "")), CDbl(Replace(vFld(Rs("codigo")), "-", "")), vFld(Rs("codigo"))) < CDbl(IIf(LCta <> "", LCta, "0")) Then
                 MsgBox1 " La cuenta " & NCta & " no debería estar después de la cuenta " & LCta & " , favor de validar plan de cuentas a copiar.", vbExclamation
                  Me.MousePointer = vbDefault
                 Exit Sub
                End If
           Else
             MsgBox1 " La cuenta " & FmtCodCuenta(vFld(Rs("Codigo"))) & " es menor 7 digitos, favor de validar plan de cuentas a copiar. ", vbExclamation
              Me.MousePointer = vbDefault
            Exit Sub
           End If

           LCta = NCta

         Rs.MoveNext

        Loop

        Call CloseRs(Rs)
'        '3360107
      
      Q1 = "INSERT INTO Cuentas ( " & FldLst & ", IdEmpresa, Ano ) SELECT " & FldLst & "," & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      Q1 = Q1 & " FROM CuentasCopy "
'      Q1 = Q1 & " WHERE CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy
      
      Rc = ExecSQL(DbMain, Q1)
      If Rc = 0 Then
         MsgBox1 "El plan de cuentas de la empresa seleccionada no tiene datos.", vbExclamation
         Me.MousePointer = vbDefault
         Exit Sub
         
      End If
#End If

   Else
         
        '3360107
'       'Chequeamos que plan de cuentas a copiar no este con conflictos
        Dim NCtaSql As String, LCtaSql As String
          
        Fld = gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano, IdPadre, IdCuenta As IdCuentaOld, IdPadre as IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII, Percepcion"
      Fld2 = " IdEmpresa, Ano, IdPadre, IdCuentaOld, IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII, Percepcion"
          
        Q1 = "SELECT " & Fld & " FROM Cuentas as Cuentas1 WHERE Cuentas1.IdEmpresa = " & IdEmpresaCopy & " AND Cuentas1.Ano = " & AnoCopy
      Q1 = Q1 & " ORDER BY Cuentas1.codigo"
        Set Rs = OpenRs(DbMain, Q1)
        Do While Rs.EOF = False

           If Len(vFld(Rs("codigo"))) > 6 Then

            NCtaSql = GetCodCta(FmtCodCuenta(vFld(Rs("Codigo"))), Nivel, nNiv, gNiveles.Largo())

                If IIf(IsNumeric(Replace(vFld(Rs("codigo")), "-", "")), CLng(Replace(vFld(Rs("codigo")), "-", "")), vFld(Rs("codigo"))) < CLng(IIf(LCtaSql <> "", LCtaSql, "0")) Then
                 MsgBox1 " La cuenta " & NCtaSql & " no debería estar después de la cuenta " & LCtaSql & " , favor de validar plan de cuentas a copiar.", vbExclamation
                  Me.MousePointer = vbDefault
                 Exit Sub
                End If
           Else
             MsgBox1 " La cuenta " & FmtCodCuenta(vFld(Rs("Codigo"))) & " es menor 7 digitos, favor de validar plan de cuentas a copiar. ", vbExclamation
              Me.MousePointer = vbDefault
            Exit Sub
           End If

           LCtaSql = NCtaSql

         Rs.MoveNext

        Loop

        Call CloseRs(Rs)
'        '3360107
  
         
         
         
      'Copiamos las Cuentas de otra empresa (IdEmpresaCopy, AnoCopy)
      'Fld = gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano, IdPadre, IdCuenta As IdCuentaOld, IdPadre as IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII, Percepcion"
      'Fld2 = " IdEmpresa, Ano, IdPadre, IdCuentaOld, IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII, Percepcion"
      Q1 = "INSERT INTO Cuentas (" & Fld2 & ") SELECT " & Fld & " FROM Cuentas as Cuentas1 WHERE Cuentas1.IdEmpresa = " & IdEmpresaCopy & " AND Cuentas1.Ano = " & AnoCopy
      Q1 = Q1 & " ORDER BY Cuentas1.IdCuenta"
      
      Rc = ExecSQL(DbMain, Q1)
      If Rc = 0 Then
         MsgBox1 "El plan de cuentas de la empresa seleccionada no tiene datos.", vbExclamation
         Exit Sub
         
      End If
      
      'actualizamos los padres
      Tbl = " Cuentas "
      sFrom = " Cuentas "
      sFrom = sFrom & " INNER JOIN Cuentas As Cuentas1 ON Cuentas.IdPadreOld = Cuentas1.IdCuentaOld "
      sFrom = sFrom & " AND Cuentas.IdEmpresa = Cuentas1.IdEmpresa AND Cuentas.Ano = Cuentas1.Ano "
      sSet = " Cuentas.IdPadre = Cuentas1.IdCuenta "
      sWhere = " WHERE Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

      'esto no funciona por el tema del Identity
'      Q1 = "INSERT INTO Cuentas ( " & FldLst & ", IdEmpresa, Ano ) SELECT " & FldLst & "," & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
'      Q1 = Q1 & " FROM Cuentas As CuentasCopy "
'      Q1 = Q1 & " WHERE CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy
   
   End If
   
   'limpiamos saldos de apertura financiero y tributario
   Q1 = "UPDATE Cuentas SET Debe = 0, Haber = 0, DebeTrib = 0, HaberTrib = 0"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Rc = ExecSQL(DbMain, Q1)
   
   'eliminamos las cuentas básicas de ParamEmpresa porque no van a servir para el nuevo plan porque los IDs son distintos
'   Q1 = "DELETE * FROM ParamEmpresa "
'   Q1 = Q1 & " WHERE Left(Tipo,3)='CTA'"
'   Rc = ExecSQL(DbMain, Q1)
   
   Q1 = " WHERE Left(Tipo,3)='CTA'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "ParamEmpresa", Q1)
   
   'Copiamos las cuentas básicas de ParamEmpresa
   
   If gEmprSeparadas Then
      Q1 = "INSERT INTO ParamEmpresa SELECT ParamEmpresaCopy.Tipo as Tipo, ParamEmpresaCopy.Codigo as Codigo, Cuentas.IdCuenta as Valor, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      
      Q1 = Q1 & " FROM (CuentasCopy INNER JOIN ParamEmpresaCopy ON CuentasCopy.idCuenta = " & SqlVal("ParamEmpresaCopy.Valor") & ")"
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo"
      Q1 = Q1 & " WHERE Left(Tipo,3)='CTA'"
      
   Else
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & " (Tipo, Codigo, Valor, IdEmpresa, Ano ) "
      Q1 = Q1 & " SELECT ParamEmpresaCopy.Tipo as Tipo, ParamEmpresaCopy.Codigo as Codigo, Cuentas.IdCuenta as  Valor,"
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      
      Q1 = Q1 & " FROM (Cuentas as CuentasCopy INNER JOIN ParamEmpresa as ParamEmpresaCopy ON CuentasCopy.idCuenta = " & SqlVal("ParamEmpresaCopy.Valor")
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = ParamEmpresaCopy.IdEmpresa AND CuentasCopy.Ano = ParamEmpresaCopy.Ano "
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy & ")"
      
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo "
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      
      Q1 = Q1 & " WHERE Left(ParamEmpresaCopy.Tipo,3)='CTA'"

   End If
   
   Rc = ExecSQL(DbMain, Q1)
      
   'eliminamos las cuentas básicas de la tabla CuentasBasicas porque no van a servir para el nuevo plan porque los IDs son distintos
'   Q1 = "DELETE * FROM CuentasBasicas "
'   Rc = ExecSQL(DbMain, Q1)
   Q1 = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "CuentasBasicas", Q1)
   
   'Copiamos las cuentas básicas de la tabla CuentasBasicas
   If gEmprSeparadas Then

#If DATACON = 1 Then       'Access

      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb", "CuentasBasicas", "CuentasBasicasCopy", , , ConnStr)
   
      Q1 = "INSERT INTO CuentasBasicas"
      Q1 = Q1 & " ( Tipo, TipoLib, TipoValor, IdCuenta, IdEmpresa, Ano )"
      Q1 = Q1 & " SELECT CuentasBasicasCopy.Tipo AS Tipo, CuentasBasicasCopy.TipoLib AS TipoLib, CuentasBasicasCopy.TipoValor AS TipoValor, Cuentas.IdCuenta AS IdCuenta, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      Q1 = Q1 & " FROM (CuentasCopy INNER JOIN CuentasBasicasCopy ON CuentasCopy.idCuenta = " & SqlVal("CuentasBasicasCopy.IdCuenta") & ") "
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo"
#End If

   Else
      Q1 = "INSERT INTO CuentasBasicas"
      Q1 = Q1 & " ( Tipo, TipoLib, TipoValor, IdCuenta, IdEmpresa, Ano )"
      Q1 = Q1 & " SELECT CuentasBasicasCopy.Tipo AS Tipo, CuentasBasicasCopy.TipoLib AS TipoLib, CuentasBasicasCopy.TipoValor AS TipoValor, Cuentas.IdCuenta AS IdCuenta, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      
      Q1 = Q1 & " FROM ( Cuentas as CuentasCopy INNER JOIN CuentasBasicas as CuentasBasicasCopy ON CuentasCopy.idCuenta = " & SqlVal("CuentasBasicasCopy.IdCuenta")
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = CuentasBasicasCopy.IdEmpresa AND CuentasCopy.Ano = CuentasBasicasCopy.Ano "
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy & ")"
      
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo "
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   
   End If
   
   Rc = ExecSQL(DbMain, Q1)
      
   'Copiamos las cuentas de ajustes extracontables Libro de Caja, de la tabla CtasAjustesExCont
   If gEmprSeparadas Then

#If DATACON = 1 Then       'Access

      Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb", "CtasAjustesExCont", "CtasAjustesExContCopy", , , ConnStr)
   
      Q1 = "INSERT INTO CtasAjustesExCont "
      Q1 = Q1 & " ( TipoAjuste, IdItem, IdCuenta, CodCuenta, IdEmpresa, Ano )"
      Q1 = Q1 & " SELECT CtasAjustesExContCopy.TipoAjuste AS TipoAjuste, CtasAjustesExContCopy.IdItem AS IdItem, Cuentas.IdCuenta AS IdCuenta, CtasAjustesExContCopy.CodCuenta AS CodCuenta, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      Q1 = Q1 & " FROM Cuentas INNER JOIN CtasAjustesExContCopy ON Cuentas.Codigo = CtasAjustesExContCopy.CodCuenta "
#End If

   Else
      Q1 = "INSERT INTO CtasAjustesExCont"
      Q1 = Q1 & " ( TipoAjuste, IdItem, IdCuenta, CodCuenta, IdEmpresa, Ano )"
      Q1 = Q1 & " SELECT CtasAjustesExContCopy.TipoAjuste AS TipoAjuste, CtasAjustesExContCopy.IdItem AS IdItem, Cuentas.IdCuenta AS IdCuenta, CtasAjustesExContCopy.CodCuenta AS CodCuenta,"
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      
      Q1 = Q1 & " FROM Cuentas INNER JOIN CtasAjustesExCont as CtasAjustesExContCopy ON Cuentas.Codigo = CtasAjustesExContCopy.CodCuenta"
      Q1 = Q1 & " AND CtasAjustesExContCopy.IdEmpresa = " & IdEmpresaCopy & " AND CtasAjustesExContCopy.Ano = " & AnoCopy
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   
   End If
   
   Rc = ExecSQL(DbMain, Q1)
      
   If gEmprSeparadas Then
      Call ExecSQL(DbMain, "Drop Table " & "ParamEmpresaCopy")
      Call ExecSQL(DbMain, "Drop Table " & "CuentasBasicasCopy")
      Call ExecSQL(DbMain, "Drop Table " & "CuentasCopy")
      Call ExecSQL(DbMain, "Drop Table " & "CtasAjustesExContCopy")
   End If
   
   
   
   Call ReadEmpresa
   
   Me.MousePointer = vbDefault
   
   lRc = vbOK
   MsgBox1 "La copia del plan de cuentas ha sido realizada exitosamente.", vbExclamation
   Unload Me
   
End Sub

Private Sub CopyConfigRemu()
   Dim Q1 As String
   Dim Rc As Long
   Dim Rs As Recordset
   Dim ExisteConfigRemu As Boolean
   Dim ConnStr As String
   Dim NConfigRemu As Integer
   Dim IdEmpresaCopy As Long, AnoCopy As Integer
   Dim fname As String
         
   If lsEmpresas.ListIndex < 0 Then
      Exit Sub
   End If
   
   IdEmpresaCopy = lsEmpresas.Matrix(MATRIX_IDEMP)
   AnoCopy = lsEmpresas.Matrix(MATRIX_MAXANO)
   
   Me.MousePointer = vbHourglass
    
   'Chequeamos que no haya una configuración previa

   Q1 = "SELECT Count(*) as Cant FROM ParamEmpresa WHERE Tipo = 'CTASREMU'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExisteConfigRemu = True
      End If
   End If
   Call CloseRs(Rs)

   If ExisteConfigRemu Then
      If MsgBox1("Esta empresa ya tiene una configuración para remuneraciones. Si copia otra configuración, perderá la actual." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
'      Call ExecSQL(DbMain, "DELETE * FROM ParamEmpresa WHERE Tipo = 'CTASREMU'")

      Q1 = "WHERE Tipo = 'CTASREMU'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

      Call DeleteSQL(DbMain, "ParamEmpresa", Q1)
      
   End If
      
   'link con la empresa seleccionada
   If gEmprSeparadas Then
   
#If DATACON = 1 Then       'Access
     '2868088
      ConnStr = "PWD=" & PASSW_PREFIX & lsEmpresas.ItemData & ";"
      'ConnStr = "PWD=" & PASSW_PREFIX_NEW & lsEmpresas.ItemData & ";"
      '2868088
      fname = gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb"
      If Not ExistFile(fname) Then
         MsgBox1 "No se ha encontrado el archivo: " & vbCrLf & vbCrLf & fname & vbCrLf & vbCrLf & "No es posible realizar la operación.", vbExclamation
         Me.MousePointer = vbDefault
         Exit Sub
      End If
         
      Call LinkMdbTable(DbMain, fname, "ParamEmpresa", "ParamEmpresaCopy", True, , ConnStr)
      Call LinkMdbTable(DbMain, fname, "Cuentas", "CuentasCopy", , , ConnStr)
      
      Q1 = "SELECT Count(*) FROM ParamEmpresaCopy WHERE Tipo = 'CTASREMU'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         NConfigRemu = vFld(Rs(0))
      End If
      Call CloseRs(Rs)
      
      Q1 = "INSERT INTO ParamEmpresa "
      Q1 = Q1 & " SELECT ParamEmpresaCopy.Tipo As Tipo, ParamEmpresaCopy.Codigo as Codigo, Cuentas.IdCuenta as Valor, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      Q1 = Q1 & " FROM (ParamEmpresaCopy INNER JOIN CuentasCopy ON " & SqlVal("ParamEmpresaCopy.Valor") & " = CuentasCopy.IdCuenta) "
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo"
      Q1 = Q1 & " WHERE ParamEmpresaCopy.Tipo = 'CTASREMU'"
      
      Rc = ExecSQL(DbMain, Q1)
#End If

   Else   'empresas juntas (SQL Server)
      Q1 = "SELECT Count(*) FROM ParamEmpresa WHERE Tipo = 'CTASREMU'"
      Q1 = Q1 & " AND ParamEmpresa.IdEmpresa = " & IdEmpresaCopy & " AND ParamEmpresa.Ano = " & AnoCopy
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         NConfigRemu = vFld(Rs(0))
      End If
      Call CloseRs(Rs)
      
      Q1 = "INSERT INTO ParamEmpresa ( Tipo, Codigo, Valor, IdEmpresa, Ano )"
      Q1 = Q1 & " SELECT ParamEmpresaCopy.Tipo As Tipo, ParamEmpresaCopy.Codigo as Codigo, Cuentas.IdCuenta as Valor, "
      Q1 = Q1 & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano "
      
      Q1 = Q1 & " FROM (ParamEmpresa As ParamEmpresaCopy INNER JOIN Cuentas as CuentasCopy ON " & SqlVal("ParamEmpresaCopy.Valor") & " = CuentasCopy.IdCuenta "
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = ParamEmpresaCopy.IdEmpresa AND CuentasCopy.Ano = ParamEmpresaCopy.Ano "
      Q1 = Q1 & " AND CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy & ")"
      
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo "
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      
      Q1 = Q1 & " WHERE ParamEmpresaCopy.Tipo = 'CTASREMU'"
      
      Rc = ExecSQL(DbMain, Q1)
      
   End If
   
   Me.MousePointer = vbDefault
   
   If Rc = 0 Then
      MsgBox1 "La Configuración de Remuneraciones de la empresa seleccionada no tiene datos o los planes de cuenta de las empresas no calzan.", vbExclamation
      Exit Sub
   ElseIf Rc < NConfigRemu Then
      MsgBox1 "Se copiaron sólo algunos registros de configuración porque hay cuentas en la empresa de origen que no están definidas en esta empresa.", vbExclamation
      Exit Sub
   End If
  
  
   If gEmprSeparadas Then
         
      Call ExecSQL(DbMain, "DROP Table " & "ParamEmpresaCopy")
      Call ExecSQL(DbMain, "DROP Table " & "CuentasCopy")
   
   End If
   
   Call ReadEmpresa
   
   Me.MousePointer = vbDefault
   
   lRc = vbOK
   MsgBox1 "La copia de la configuración de remnueraciones ha finalizado.", vbExclamation
   Unload Me
   
End Sub

Private Sub bt_Copy_Click()

'   If MsgBox1("Una vez realizada la copia no podrá volver a la configuración actual." & vbCrLf & vbCrLf & "¿Está seguro que desea copiar la configuración de esta empresa?", vbQuestion + vbYesNoCancel) <> vbYes Then
'      Exit Sub
'   End If

   If lCopyObj = C_COPYPLAN Then
      Call CopyPlan
   ElseIf lCopyObj = C_COPYCONFIGREMU Then
      Call CopyConfigRemu
   Else
      CopyConfigImpAdic
   End If
   
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   
   If lCopyObj = C_COPYPLAN Then
      Me.Caption = "Copiar Plan de Cuentas de otra empresa"
   ElseIf lCopyObj = C_COPYCONFIGREMU Then
      Me.Caption = "Copiar configuración de Remuneraciones de otra empresa"
   Else
      Me.Caption = "Copiar configuración de Impuestos Adicionales de otra empresa"
   End If

   Call FillList
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv
   
End Sub
Private Sub FillList()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
    
   Set lsEmpresas = New ClsCombo
   Call lsEmpresas.SetControl(Ls_Empresas)
   
'   If gUsuario.Nombre <> gAdmUser Then
'      Wh = " AND id IN(" & gUsuario.idEmpresas & ")"
'   End If
   
   Q1 = "SELECT DISTINCT Max(Ano) as MaxAno, Empresas.idEmpresa, Rut, NombreCorto "
   Q1 = Q1 & " FROM Empresas"
   Q1 = Q1 & " INNER JOIN EmpresasAno ON EmpresasAno.idEmpresa = Empresas.idEmpresa"
   Q1 = Q1 & " WHERE Empresas.idEmpresa <>" & gEmpresa.id
   Q1 = Q1 & Wh
   Q1 = Q1 & " GROUP BY Empresas.idEmpresa,Rut,NombreCorto"
   Q1 = Q1 & " ORDER BY NombreCorto"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
      lsEmpresas.AddItem FmtStRut(vFld(Rs("Rut"))) & "     " & vFld(Rs("NombreCorto"))
      lsEmpresas.ItemData(lsEmpresas.NewIndex) = vFld(Rs("Rut"))
      lsEmpresas.Matrix(MATRIX_MAXANO, lsEmpresas.NewIndex) = vFld(Rs("MaxAno"))
      lsEmpresas.Matrix(MATRIX_IDEMP, lsEmpresas.NewIndex) = vFld(Rs("idEmpresa"))
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
End Sub

Public Function FCopy() As Integer
   lCopyObj = C_COPYPLAN
   
   Me.Show vbModal
   
   FCopy = lRc

End Function
Public Function FCopyConfigRemu() As Integer
   lCopyObj = C_COPYCONFIGREMU
   
   Me.Show vbModal
   
   FCopyConfigRemu = lRc

End Function
Public Function FCopyConfigImpAdic() As Integer
   lCopyObj = C_COPYCONFIGIMPADIC
   
   Me.Show vbModal
   
   FCopyConfigImpAdic = lRc

End Function

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   
End Function

Private Sub CopyConfigImpAdic()
   Dim Q1 As String
   Dim Rc As Long
   Dim Rs As Recordset
   Dim ExisteConfigImpAdic As Boolean
   Dim ConnStr As String
   Dim NImpAdic As Integer
   Dim IdEmpresaCopy As Long, AnoCopy As Integer
   Dim fname As String
         
   If lsEmpresas.ListIndex < 0 Then
      Exit Sub
   End If
    
   IdEmpresaCopy = lsEmpresas.Matrix(MATRIX_IDEMP)
   AnoCopy = lsEmpresas.Matrix(MATRIX_MAXANO)
   
   'Chequeamos que no haya una configuración previa

   Q1 = "SELECT Count(*) as Cant FROM ImpAdic "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If vFld(Rs("Cant")) > 0 Then
         ExisteConfigImpAdic = True
      End If
   End If
   Call CloseRs(Rs)
   
   If ExisteConfigImpAdic Then
      If MsgBox1("Esta empresa ya tiene una configuración para impuestos adicionales. Si copia otra configuración, perderá la actual." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
'      Call ExecSQL(DbMain, "DELETE * FROM ImpAdic")
      Call DeleteSQL(DbMain, "ImpAdic", " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   End If
   
   Me.MousePointer = vbHourglass
     
   'link con la empresa seleccionada
   
   If gEmprSeparadas Then
   
#If DATACON = 1 Then       'Access
    
    '2868088
      ConnStr = "PWD=" & PASSW_PREFIX & lsEmpresas.ItemData & ";"
      'ConnStr = "PWD=" & PASSW_PREFIX_NEW & lsEmpresas.ItemData & ";"
      '2868088
      
      fname = gDbPath & "\Empresas\" & AnoCopy & "\" & lsEmpresas.ItemData & ".mdb"
      If Not ExistFile(fname) Then
         MsgBox1 "No se ha encontrado el archivo: " & vbCrLf & vbCrLf & fname & vbCrLf & vbCrLf & "No es posible realizar la operación.", vbExclamation
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      
      Call LinkMdbTable(DbMain, fname, "ImpAdic", "ImpAdicCopy", , , ConnStr)
      Call LinkMdbTable(DbMain, fname, "Cuentas", "CuentasCopy", , , ConnStr)
      
      Q1 = "SELECT Count(*) FROM ImpAdicCopy"
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         NImpAdic = vFld(Rs(0))
      End If
      Call CloseRs(Rs)
      
      Q1 = "INSERT INTO ImpAdic "
      Q1 = Q1 & " ( TipoLib, TipoValor, IdCuenta, Tasa, "
      Q1 = Q1 & " EsRecuperable, IdEmpresa , Ano ) "
      Q1 = Q1 & " SELECT ImpAdicCopy.TipoLib As TipoLib, ImpAdicCopy.TipoValor as TipoValor, Cuentas.IdCuenta as IdCuenta, ImpAdicCopy.Tasa as Tasa, "
      Q1 = Q1 & " ImpAdicCopy.EsRecuperable as EsRecuperable, " & gEmpresa.id & " as IdEmpresa, " & gEmpresa.Ano & " as Ano"
      Q1 = Q1 & " FROM (ImpAdicCopy LEFT JOIN CuentasCopy ON ImpAdicCopy.IdCuenta = CuentasCopy.IdCuenta) "
      Q1 = Q1 & " LEFT JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo"
      
      Rc = ExecSQL(DbMain, Q1)
#End If

   Else
      Q1 = "SELECT Count(*) FROM ImpAdic"
      Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaCopy & " AND Ano = " & AnoCopy
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         NImpAdic = vFld(Rs(0))
      End If
      Call CloseRs(Rs)
      
      If gDbType = SQL_ACCESS Then
      
         Q1 = "INSERT INTO ImpAdic "
         Q1 = Q1 & " ( TipoLib, TipoValor, IdCuenta, Tasa, "
         Q1 = Q1 & " EsRecuperable, IdEmpresa, Ano ) "
         Q1 = Q1 & " SELECT ImpAdicCopy.TipoLib As TipoLib, ImpAdicCopy.TipoValor as TipoValor, Cuentas.IdCuenta as IdCuenta, ImpAdicCopy.Tasa as Tasa, "
         Q1 = Q1 & " ImpAdicCopy.EsRecuperable as EsRecuperable, " & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano "
         Q1 = Q1 & " FROM (ImpAdicCopy LEFT JOIN CuentasCopy ON ImpAdicCopy.IdCuenta = CuentasCopy.IdCuenta) "
         Q1 = Q1 & " LEFT JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo"
      
      Else    'SQL Server o MySQL (esto es porque primero los junta y después los corta y acá podemos poner restricciones en el JOIN
      
         Q1 = "INSERT INTO ImpAdic "
         Q1 = Q1 & " ( TipoLib, TipoValor, IdCuenta, Tasa, "
         Q1 = Q1 & " EsRecuperable, IdEmpresa, Ano ) "
         Q1 = Q1 & " SELECT ImpAdicCopy.TipoLib As TipoLib, ImpAdicCopy.TipoValor as TipoValor, Cuentas.IdCuenta as IdCuenta, ImpAdicCopy.Tasa as Tasa, "
         Q1 = Q1 & " ImpAdicCopy.EsRecuperable as EsRecuperable, " & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano "
         
         Q1 = Q1 & " FROM ( ImpAdic as ImpAdicCopy LEFT JOIN Cuentas as CuentasCopy ON ImpAdicCopy.IdCuenta = CuentasCopy.IdCuenta "
         Q1 = Q1 & " AND ImpAdicCopy.IdEmpresa = CuentasCopy.IdEmpresa AND ImpAdicCopy.Ano = CuentasCopy.Ano  "
         Q1 = Q1 & " AND CuentasCopy.IdEmpresa = " & IdEmpresaCopy & " AND CuentasCopy.Ano = " & AnoCopy & ")"
         
         Q1 = Q1 & " LEFT JOIN Cuentas ON CuentasCopy.Codigo = Cuentas.Codigo "
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      
      End If
      
      Rc = ExecSQL(DbMain, Q1)
   
   End If
   Me.MousePointer = vbDefault
   
   If Rc = 0 Then
      MsgBox1 "La Configuración de Impuestos Adicionales de la empresa seleccionada no tiene datos o los planes de cuenta de las empresas no calzan.", vbExclamation
      Exit Sub
   ElseIf Rc < NImpAdic Then
      MsgBox1 "Se copiaron sólo algunos registros de configuración porque hay cuentas en la empresa de origen que no están definidas en esta empresa.", vbExclamation
      Exit Sub
   End If
         
   If gEmprSeparadas Then
      Call ExecSQL(DbMain, "Drop Table " & "ImpAdicCopy")
      Call ExecSQL(DbMain, "Drop Table " & "CuentasCopy")
   End If
   
   lRc = vbOK
   MsgBox1 "La copia de la configuración de impuestos adicionales ha finalizado." & vbCrLf & vbCrLf & "Es posible que falten algunas cuentas de configuración, porque pueden haber algunas cuentas en la empresa de origen que no estén definidas en esta empresa.", vbExclamation
   Unload Me
   
End Sub

Private Sub Ls_Empresas_DblClick()
Call bt_Copy_Click
End Sub
