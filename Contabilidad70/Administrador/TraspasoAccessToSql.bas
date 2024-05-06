Attribute VB_Name = "TraspasoAccessToSql"
Private VerDBAccess As Long
Private IdEmpresaTras As Long
Private i, j As Integer

Public Sub TrasLpContab(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer, GIdEmpresaTras As Long, Optional ByVal PG As ProgressBar, Optional ByVal Txt As Label)
    
    j = 75
    i = 0
    PG.Max = j
    PG.Value = i
    
    MousePointer = vbHourglass
    DoEvents

        

'    i = i + 1
'    PG.Value = i
'    Txt.Refresh
'    Txt.Caption = "Proceso de Traspaso... Año " & ano & " (" & PG.Value & " / " & PG.Max & ")"
'    Call CorrigeBaseAdmSQLServer 'NO VA ESTE CORRIGE BASE
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call CorrigeBaseSQLServer
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasCapPropioSimplAnual(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasCodActiv(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasControlEmpresa(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasEquivalencia(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasFactorActAnual(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasIFRS_PlanIFRS(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasImpuestos(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasIPC(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasMonedas(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasParam(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasPerfiles(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasPlanAvanzado(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasPlanBasico(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasPlanCuentasSII(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasPlanIntermedio(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasRazonesFin(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasRegiones(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasTimbraje(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasTipoDocs(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasUsuarioEmpresa(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasUsuarios(DBSql, DbAccess)
'    PG.Value = i
'    Txt.Caption = "Proceso de Traspaso... Año " & ano & " (" & PG.Value & " / " & PG.Max & ")"
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    Call TrasTipoValor(DBSql, DbAccess)
    i = i + 1
    PG.Value = i
    Txt.Refresh
    Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
    'Call TrasParamEmpresa(DBSql, DbAccess, IdEmpresa, Ano)
    IdEmpresaTras = GIdEmpresaTras
    Call TrasEmpresasAno(DBSql, DbAccess, IdEmpresa, Ano)

End Sub

Public Sub TrasLPEmpresa(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer, GIdEmpresaTras As Long, Optional ByVal PG As ProgressBar, Optional ByVal Txt As Label)
 
    'Call TrasEmpresas(DBSql, DbAccess)
    Dim Rs As dao.Recordset
    MousePointer = vbHourglass
    DoEvents
    
     'ffv 3384570
        Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
        'Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
     'ffv 3384570
   Set Rs = OpenRsDao(DbAccess, Q1)
   If Rs.EOF = False Then
      VerDBAccess = Val(vFldDao(Rs("Valor")))
   End If
   Call CloseRs(Rs)
    

'Call TrasEmpresasAno(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasParamEmpresa(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasEntidades(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCuentas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasComprobante(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasAreaNegocio(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCentroCosto(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCartola(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasSucursales(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDocumento(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasMovDocumento(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasMovComprobante(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCT_Comprobante(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCT_MovComprobante(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
'Call TrasCT_ComprobanteBase(DBSql, DBAccess, IdEmpresa, ano) 'no va ya que esta tabla existe solo en SQL por ende no viene de Access
Call TrasAjusteIVAMensual(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasAjustesExtLibCaja(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasAFGrupos(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasAFComponentes(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasMovActivoFijo(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasActFijoFicha(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasActFijoCompsFicha(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasAsistImpPrimCat(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasBaseImponible14D(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasBaseImponible14ter(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasColores(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasContactos(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCtasAjustesExCont(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCtasAjustesExContRLI(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCuentasBasicas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasCuentasRazon(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDetCapPropioSimpl(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDetCartola(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasPercepciones(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDetPercepciones(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDetSaldosAp(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasLibroCaja(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasDocCuotas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasEmpresa(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasEstadoMes(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasFirmas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasGlosas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasImpAdic(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasInfoAnualDJ1847(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasLockAction(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasLogComprobantes(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasMembrete(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasNotas(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasParamRazon(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasPropIVA_TotMensual(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
Call TrasSocios(DBSql, DbAccess, IdEmpresa, Ano)
i = i + 1
PG.Value = i
Txt.Refresh
Txt.Caption = "Proceso de Traspaso... Año " & Ano & " (" & PG.Value & " / " & PG.Max & ")"
'Call TrasTipoValor(DBSql, DbAccess, IdEmpresa, Ano)


Txt.Caption = "Proceso de Traspaso... Año " & Ano & " Terminado"


End Sub




Public Sub TrasEmpresas(DBSql As ADODB.Connection, DbAccess As Database)
   
   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Empresas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Empresas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)
    
    
'   Q1 = "DELETE FROM EMPRESAS "
'   Set Rs1 = OpenRs(DBSql, Q1)

   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,Rut"
   Q1 = Q1 & " ,NombreCorto"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,RutDisp"
   'Q1 = Q1 & " ,Import2"
   Q1 = Q1 & " FROM Empresas "
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Empresas"
            Q1 = Q1 & " WHERE Trim(NombreCorto) = '" & Trim(vFldDao(Rs("NombreCorto"))) & "'"
            'Q1 = Q1 & " WHERE IdTras = " & vFldDao(Rs("IdEmpresa"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
   
                'Q1 = " SET IDENTITY_INSERT Empresas ON "
                Q1 = " INSERT INTO Empresas"
                Q1 = Q1 & " (Rut"
                Q1 = Q1 & " ,NombreCorto"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,RutDisp"
                Q1 = Q1 & " ,IdTras)"
                'Q1 = Q1 & " ,Import2)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("Rut")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombreCorto")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("RutDisp")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdEmpresa")) & ")"
                'Q1 = Q1 & " ," & vFldDao(Rs("Import2")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Empresas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Empresas"
                Q1 = Q1 & " SET Rut = '" & vFldDao(Rs("Rut")) & "'"
                Q1 = Q1 & " ,NombreCorto = '" & vFldDao(Rs("NombreCorto")) & "'"
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,RutDisp = '" & vFldDao(Rs("RutDisp")) & "'"
                Q1 = Q1 & " ,IdTras = '" & vFldDao(Rs("IdEmpresa")) & "'"
                'Q1 = Q1 & " ,Import2 = " & vFldDao(Rs("Import2"))
                'Q1 = Q1 & " WHERE IdTras = " & vFldDao(Rs("IdEmpresa"))
                Q1 = Q1 & " WHERE Trim(NombreCorto) = '" & Trim(vFldDao(Rs("NombreCorto"))) & "'"
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasEmpresa(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)
   
   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
    
    
   Q1 = "DELETE FROM EMPRESA WHERE ID = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs1 = OpenRs(DBSql, Q1)

    Q1 = "SELECT Id"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,Rut"
    Q1 = Q1 & " ,NombreCorto"
    Q1 = Q1 & " ,RazonSocial"
    Q1 = Q1 & " ,ApPaterno"
    Q1 = Q1 & " ,ApMaterno"
    Q1 = Q1 & " ,Nombre"
    Q1 = Q1 & " ,Calle"
    Q1 = Q1 & " ,Numero"
    Q1 = Q1 & " ,Dpto"
    Q1 = Q1 & " ,Telefonos"
    Q1 = Q1 & " ,Fax"
    Q1 = Q1 & " ,Region"
    Q1 = Q1 & " ,Comuna"
    Q1 = Q1 & " ,Ciudad"
    Q1 = Q1 & " ,Giro"
    Q1 = Q1 & " ,ActEconom"
    Q1 = Q1 & " ,CodActEconom"
    Q1 = Q1 & " ,DomPostal"
    Q1 = Q1 & " ,ComunaPostal"
    Q1 = Q1 & " ,Email"
    Q1 = Q1 & " ,Web"
    Q1 = Q1 & " ,FechaConstitucion"
    Q1 = Q1 & " ,FechaInicioAct"
    Q1 = Q1 & " ,RepConjunta"
    Q1 = Q1 & " ,RutRepLegal1"
    Q1 = Q1 & " ,RepLegal1"
    Q1 = Q1 & " ,RutRepLegal2"
    Q1 = Q1 & " ,RepLegal2"
    Q1 = Q1 & " ,Contador"
    Q1 = Q1 & " ,RutContador"
    Q1 = Q1 & " ,TipoContrib"
    Q1 = Q1 & " ,TransaBolsa"
    Q1 = Q1 & " ,Franq14bis"
    Q1 = Q1 & " ,FranqLey18392"
    Q1 = Q1 & " ,FranqDL600"
    Q1 = Q1 & " ,FranqDL701"
    Q1 = Q1 & " ,FranqDS341"
    Q1 = Q1 & " ,Opciones"
    Q1 = Q1 & " ,TContribFUT"
    Q1 = Q1 & " ,Franq14ter"
    Q1 = Q1 & " ,Franq14quater"
    Q1 = Q1 & " ,ObligaLibComprasVentas"
    Q1 = Q1 & " ,FranqRentaAtribuida"
    Q1 = Q1 & " ,FranqSemiIntegrado"
    Q1 = Q1 & " ,FranqSocProfPrimCat"
    Q1 = Q1 & " ,FranqSocProfSegCat"
    Q1 = Q1 & " ,Franq14ASemiIntegrado"
    Q1 = Q1 & " ,FranqProPymeGeneral"
    Q1 = Q1 & " ,FranqProPymeTransp"
    Q1 = Q1 & " ,FranqRentasPresuntas"
    Q1 = Q1 & " ,FranqRentaEfectiva"
    Q1 = Q1 & " ,FranqOtro"
    Q1 = Q1 & " ,FranqNoSujetoArt14"
    Q1 = Q1 & " ,CodArea"
    Q1 = Q1 & " ,Celular"
    Q1 = Q1 & " ,Villa"
    Q1 = Q1 & " From Empresa"
    Q1 = Q1 & " WHERE Id = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
   
           Q1 = " INSERT INTO Empresa"
           Q1 = Q1 & " (Id"
           Q1 = Q1 & " ,Ano"
           Q1 = Q1 & " ,Rut"
           Q1 = Q1 & " ,NombreCorto"
           Q1 = Q1 & " ,RazonSocial"
           Q1 = Q1 & " ,ApPaterno"
           Q1 = Q1 & " ,ApMaterno"
           Q1 = Q1 & " ,Nombre"
           Q1 = Q1 & " ,Calle"
           Q1 = Q1 & " ,Numero"
           Q1 = Q1 & " ,Dpto"
           Q1 = Q1 & " ,Telefonos"
           Q1 = Q1 & " ,Fax"
           Q1 = Q1 & " ,Region"
           Q1 = Q1 & " ,Comuna"
           Q1 = Q1 & " ,Ciudad"
           Q1 = Q1 & " ,Giro"
           Q1 = Q1 & " ,ActEconom"
           Q1 = Q1 & " ,CodActEconom"
           Q1 = Q1 & " ,DomPostal"
           Q1 = Q1 & " ,ComunaPostal"
           Q1 = Q1 & " ,Email"
           Q1 = Q1 & " ,Web"
           Q1 = Q1 & " ,FechaConstitucion"
           Q1 = Q1 & " ,FechaInicioAct"
           Q1 = Q1 & " ,RepConjunta"
           Q1 = Q1 & " ,RutRepLegal1"
           Q1 = Q1 & " ,RepLegal1"
           Q1 = Q1 & " ,RutRepLegal2"
           Q1 = Q1 & " ,RepLegal2"
           Q1 = Q1 & " ,Contador"
           Q1 = Q1 & " ,RutContador"
           Q1 = Q1 & " ,TipoContrib"
           Q1 = Q1 & " ,TransaBolsa"
           Q1 = Q1 & " ,Franq14bis"
           Q1 = Q1 & " ,FranqLey18392"
           Q1 = Q1 & " ,FranqDL600"
           Q1 = Q1 & " ,FranqDL701"
           Q1 = Q1 & " ,FranqDS341"
           Q1 = Q1 & " ,Opciones"
           Q1 = Q1 & " ,TContribFUT"
           Q1 = Q1 & " ,Franq14ter"
           Q1 = Q1 & " ,Franq14quater"
           Q1 = Q1 & " ,ObligaLibComprasVentas"
           Q1 = Q1 & " ,FranqRentaAtribuida"
           Q1 = Q1 & " ,FranqSemiIntegrado"
           Q1 = Q1 & " ,FranqSocProfPrimCat"
           Q1 = Q1 & " ,FranqSocProfSegCat"
           Q1 = Q1 & " ,Franq14ASemiIntegrado"
           Q1 = Q1 & " ,FranqProPymeGeneral"
           Q1 = Q1 & " ,FranqProPymeTransp"
           Q1 = Q1 & " ,FranqRentasPresuntas"
           Q1 = Q1 & " ,FranqRentaEfectiva"
           Q1 = Q1 & " ,FranqOtro"
           Q1 = Q1 & " ,FranqNoSujetoArt14"
           Q1 = Q1 & " ,CodArea"
           Q1 = Q1 & " ,Celular"
           Q1 = Q1 & " ,Villa)"
           Q1 = Q1 & " Values"
           Q1 = Q1 & " (" & IdEmpresa
           Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("Rut")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("NombreCorto")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("RazonSocial")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("ApPaterno")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("ApMaterno")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Calle")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Numero")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Dpto")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Telefonos")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Fax")) & "'"
           Q1 = Q1 & " ," & vFldDao(Rs("Region"))
           Q1 = Q1 & " ," & vFldDao(Rs("Comuna"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("Ciudad")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Giro")) & "'"
           Q1 = Q1 & " ," & vFldDao(Rs("ActEconom"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("CodActEconom")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("DomPostal")) & "'"
           Q1 = Q1 & " ," & vFldDao(Rs("ComunaPostal"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("Email")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Web")) & "'"
           Q1 = Q1 & " ," & vFldDao(Rs("FechaConstitucion"))
           Q1 = Q1 & " ," & vFldDao(Rs("FechaInicioAct"))
           Q1 = Q1 & " ," & vFldDao(Rs("RepConjunta"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("RutRepLegal1")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("RepLegal1")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("RutRepLegal2")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("RepLegal2")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("Contador")) & "'"
           Q1 = Q1 & " ,'" & vFldDao(Rs("RutContador")) & "'"
           Q1 = Q1 & " ," & vFldDao(Rs("TipoContrib"))
           Q1 = Q1 & " ," & vFldDao(Rs("TransaBolsa"))
           Q1 = Q1 & " ," & vFldDao(Rs("Franq14bis"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqLey18392"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqDL600"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqDL701"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqDS341"))
           Q1 = Q1 & " ," & vFldDao(Rs("Opciones"))
           Q1 = Q1 & " ," & vFldDao(Rs("TContribFUT"))
           Q1 = Q1 & " ," & vFldDao(Rs("Franq14ter"))
           Q1 = Q1 & " ," & vFldDao(Rs("Franq14quater"))
           Q1 = Q1 & " ," & vFldDao(Rs("ObligaLibComprasVentas"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqRentaAtribuida"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqSemiIntegrado"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqSocProfPrimCat"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqSocProfSegCat"))
           Q1 = Q1 & " ," & vFldDao(Rs("Franq14ASemiIntegrado"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqProPymeGeneral"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqProPymeTransp"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqRentasPresuntas"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqRentaEfectiva"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqOtro"))
           Q1 = Q1 & " ," & vFldDao(Rs("FranqNoSujetoArt14"))
           Q1 = Q1 & " ," & vFldDao(Rs("CodArea"))
           Q1 = Q1 & " ," & vFldDao(Rs("Celular"))
           Q1 = Q1 & " ,'" & vFldDao(Rs("Villa")) & "')"
           'Q1 = Q1 & " SET IDENTITY_INSERT Empresas OFF  "
           Call ExecSQL(DBSql, Q1)

      Rs.MoveNext
   Loop
   Call CloseRs(Rs)


End Sub

Public Sub ModCtaParamEmpresa(DBSql As ADODB.Connection, IdEmpresa As Long, Ano As Integer)
   Dim Q1 As String
     
     
    'Modifica el valor en paramempresa cuando son cuentas las regulariza con las cuentas actuales
    Q1 = " UPDATE P"
    Q1 = Q1 & " SET P.Valor = ISNULL(C.idCuenta,'0')"
    Q1 = Q1 & " FROM ParamEmpresa AS P"
    Q1 = Q1 & " LEFT JOIN Cuentas  as C ON convert(varchar,C.IdTras) = P.valor AND C.Ano = P.Ano AND C.IdEmpresa = C.IdEmpresa"
    Q1 = Q1 & " WHERE TIPO LIKE '%CTA%' AND TIPO <> 'PLANCTAS'"
    Q1 = Q1 & " AND P.IdEmpresa =" & IdEmpresa & " AND P.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)
    
''    'Actualiza las cuentas padres con las cuentas actuales se separon por niveles
'     Q1 = " UPDATE CU"
'     Q1 = Q1 & " Set CU.idPadre = CUE.Idcuenta - 1"
'     Q1 = Q1 & " FROM CUENTAS CU INNER JOIN (SELECT idPadre, Min(idCuenta) Idcuenta"
'     Q1 = Q1 & "      From CUENTAS"
'     Q1 = Q1 & "      Where Ano = " & Ano
'     Q1 = Q1 & "      AND IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & "      AND Nivel = 2"
'     Q1 = Q1 & "      GROUP BY idPadre"
'     Q1 = Q1 & " ) CUE"
'     Q1 = Q1 & " ON CU.idPadre = CUE.idPadre"
'     Q1 = Q1 & " Where CU.Ano = " & Ano
'     Q1 = Q1 & " AND CU.IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & " AND CU.Nivel = 2"
'     Call ExecSQL(DBSql, Q1)
'
'     Q1 = " UPDATE CU"
'     Q1 = Q1 & " Set CU.idPadre = CUE.Idcuenta - 1"
'     Q1 = Q1 & " FROM CUENTAS CU INNER JOIN (SELECT idPadre, Min(idCuenta) Idcuenta"
'     Q1 = Q1 & "      From CUENTAS"
'     Q1 = Q1 & "      Where Ano = " & Ano
'     Q1 = Q1 & "      AND IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & "      AND Nivel = 3"
'     Q1 = Q1 & "      GROUP BY idPadre"
'     Q1 = Q1 & " ) CUE"
'     Q1 = Q1 & " ON CU.idPadre = CUE.idPadre"
'     Q1 = Q1 & " Where CU.Ano = " & Ano
'     Q1 = Q1 & " AND CU.IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & " AND CU.Nivel = 3"
'     Call ExecSQL(DBSql, Q1)
'
'     Q1 = " UPDATE CU"
'     Q1 = Q1 & " Set CU.idPadre = CUE.Idcuenta - 1"
'     Q1 = Q1 & " FROM CUENTAS CU INNER JOIN (SELECT idPadre, Min(idCuenta) Idcuenta"
'     Q1 = Q1 & "      From CUENTAS"
'     Q1 = Q1 & "      Where Ano = " & Ano
'     Q1 = Q1 & "      AND IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & "      AND Nivel = 4"
'     Q1 = Q1 & "      GROUP BY idPadre"
'     Q1 = Q1 & " ) CUE"
'     Q1 = Q1 & " ON CU.idPadre = CUE.idPadre"
'     Q1 = Q1 & " Where CU.Ano = " & Ano
'     Q1 = Q1 & " AND CU.IdEmpresa = " & IdEmpresa
'     Q1 = Q1 & " AND CU.Nivel = 4"
'     Call ExecSQL(DBSql, Q1)
    
                

End Sub

Public Sub TrasEntidades(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
   'On Error Resume Next
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Entidades' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Entidades ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)


   Q1 = "SELECT IdEntidad"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Rut"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Direccion"
   Q1 = Q1 & " ,Region"
   Q1 = Q1 & " ,Comuna"
   Q1 = Q1 & " ,Ciudad"
   Q1 = Q1 & " ,Telefonos"
   Q1 = Q1 & " ,Fax"
   Q1 = Q1 & " ,ActEcon"
   Q1 = Q1 & " ,CodActEcon"
   Q1 = Q1 & " ,DomPostal"
   Q1 = Q1 & " ,ComPostal"
   Q1 = Q1 & " ,Email"
   Q1 = Q1 & " ,Web"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Obs"
   Q1 = Q1 & " ,Clasif0"
   Q1 = Q1 & " ,Clasif1"
   Q1 = Q1 & " ,Clasif2"
   Q1 = Q1 & " ,Clasif3"
   Q1 = Q1 & " ,Clasif4"
   Q1 = Q1 & " ,Clasif5"
   Q1 = Q1 & " ,Giro"
   Q1 = Q1 & " ,NotValidRut"
   Q1 = Q1 & " ,EsSupermercado"
   Q1 = Q1 & " ,EntRelacionada"
   Q1 = Q1 & " ,CodCtaAfecto"
   Q1 = Q1 & " ,CodCtaExento"
   Q1 = Q1 & " ,CodCtaTotal"
   Q1 = Q1 & " ,PropIVA"
   Q1 = Q1 & " ,CodCCostoAfecto"
   Q1 = Q1 & " ,CodAreaNegAfecto"
   Q1 = Q1 & " ,CodCCostoExento"
   Q1 = Q1 & " ,CodAreaNegExento"
   Q1 = Q1 & " ,CodCCostoTotal"
   Q1 = Q1 & " ,CodAreaNegTotal"
   Q1 = Q1 & " ,CodCtaAfectoVta"
   Q1 = Q1 & " ,CodCtaExentoVta"
   Q1 = Q1 & " ,CodCtaTotalVta"
   Q1 = Q1 & " ,CodCCostoAfectoVta"
   Q1 = Q1 & " ,CodAreaNegAfectoVta"
   Q1 = Q1 & " ,CodCCostoExentoVta"
   Q1 = Q1 & " ,CodAreaNegExentoVta"
   Q1 = Q1 & " ,CodCCostoTotalVta"
   Q1 = Q1 & " ,CodAreaNegTotalVta"
   Q1 = Q1 & " ,EsDelGiro"
   Q1 = Q1 & " ,FranqTribEnt"
   Q1 = Q1 & " ,Ret3Porc"
   Q1 = Q1 & " ,FDesde3Porc"
   Q1 = Q1 & " ,FHasta3Porc"
   Q1 = Q1 & " FROM Entidades "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
            If vFldDao(Rs("Rut")) = "A-28120087" Then
            Q1 = ""
            End If
            If IsNumeric(vFldDao(Rs("Rut"))) Then
            
            Q1 = "SELECT * "
            Q1 = Q1 & " From Entidades"
            Q1 = Q1 & " WHERE Rut = '" & Replace(Replace(vFldDao(Rs("Rut")), "-", ""), ".", "") & "'"
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdEntidad"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
                If Rs1.EOF = True Then
                
                    'Q1 = " SET IDENTITY_INSERT Entidades ON "
                    Q1 = " INSERT INTO Entidades"
                    Q1 = Q1 & " (IdEmpresa"
                    Q1 = Q1 & " ,Rut"
                    Q1 = Q1 & " ,Codigo"
                    Q1 = Q1 & " ,Nombre"
                    Q1 = Q1 & " ,Direccion"
                    Q1 = Q1 & " ,Region"
                    Q1 = Q1 & " ,Comuna"
                    Q1 = Q1 & " ,Ciudad"
                    Q1 = Q1 & " ,Telefonos"
                    Q1 = Q1 & " ,Fax"
                    Q1 = Q1 & " ,ActEcon"
                    Q1 = Q1 & " ,CodActEcon"
                    Q1 = Q1 & " ,DomPostal"
                    Q1 = Q1 & " ,ComPostal"
                    Q1 = Q1 & " ,Email"
                    Q1 = Q1 & " ,Web"
                    Q1 = Q1 & " ,Estado"
                    Q1 = Q1 & " ,Obs"
                    Q1 = Q1 & " ,Clasif0"
                    Q1 = Q1 & " ,Clasif1"
                    Q1 = Q1 & " ,Clasif2"
                    Q1 = Q1 & " ,Clasif3"
                    Q1 = Q1 & " ,Clasif4"
                    Q1 = Q1 & " ,Clasif5"
                    Q1 = Q1 & " ,Giro"
                    Q1 = Q1 & " ,NotValidRut"
                    Q1 = Q1 & " ,EsSupermercado"
                    Q1 = Q1 & " ,EntRelacionada"
                    Q1 = Q1 & " ,CodCtaAfecto"
                    Q1 = Q1 & " ,CodCtaExento"
                    Q1 = Q1 & " ,CodCtaTotal"
                    Q1 = Q1 & " ,PropIVA"
                    Q1 = Q1 & " ,CodCCostoAfecto"
                    Q1 = Q1 & " ,CodAreaNegAfecto"
                    Q1 = Q1 & " ,CodCCostoExento"
                    Q1 = Q1 & " ,CodAreaNegExento"
                    Q1 = Q1 & " ,CodCCostoTotal"
                    Q1 = Q1 & " ,CodAreaNegTotal"
                    Q1 = Q1 & " ,CodCtaAfectoVta"
                    Q1 = Q1 & " ,CodCtaExentoVta"
                    Q1 = Q1 & " ,CodCtaTotalVta"
                    Q1 = Q1 & " ,CodCCostoAfectoVta"
                    Q1 = Q1 & " ,CodAreaNegAfectoVta"
                    Q1 = Q1 & " ,CodCCostoExentoVta"
                    Q1 = Q1 & " ,CodAreaNegExentoVta"
                    Q1 = Q1 & " ,CodCCostoTotalVta"
                    Q1 = Q1 & " ,CodAreaNegTotalVta"
                    Q1 = Q1 & " ,EsDelGiro"
                    Q1 = Q1 & " ,FranqTribEnt"
                    Q1 = Q1 & " ,Ret3Porc"
                    Q1 = Q1 & " ,FDesde3Porc"
                    Q1 = Q1 & " ,FHasta3Porc"
                    Q1 = Q1 & " ,IdTras)"
                    Q1 = Q1 & " Values"
                    Q1 = Q1 & " ('" & IdEmpresa & "'"
                    Q1 = Q1 & " ,'" & Replace(Replace(vFldDao(Rs("Rut")), "-", ""), ".", "") & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                    Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Nombre")), Chr(39), "") & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Direccion")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("Region"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Comuna"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Ciudad")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Telefonos")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Fax")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("ActEcon"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodActEcon")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("DomPostal")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("ComPostal"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Email")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Web")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Obs")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif0"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif1"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif2"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif3"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif4"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasif5"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Giro")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("NotValidRut"))
                    Q1 = Q1 & " ," & vFldDao(Rs("EsSupermercado"))
                    Q1 = Q1 & " ," & vFldDao(Rs("EntRelacionada"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaAfecto")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaExento")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaTotal")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("PropIVA"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoAfecto")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegAfecto")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoExento")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegExento")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoTotal")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegTotal")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaAfectoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaExentoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaTotalVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoAfectoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegAfectoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoExentoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegExentoVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCCostoTotalVta")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodAreaNegTotalVta")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("EsDelGiro"))
                    Q1 = Q1 & " ," & vFldDao(Rs("FranqTribEnt"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Ret3Porc"))
                    Q1 = Q1 & " ," & vFldDao(Rs("FDesde3Porc"))
                    Q1 = Q1 & " ," & vFldDao(Rs("FHasta3Porc"))
                    Q1 = Q1 & " ," & vFldDao(Rs("IdEntidad")) & ")"
                    'Q1 = Q1 & " SET IDENTITY_INSERT Entidades OFF  "
                    Call ExecSQL(DBSql, Q1)
                    
                Else
                
                    Q1 = " UPDATE Entidades"
                    Q1 = Q1 & " SET Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                    Q1 = Q1 & " ,Nombre = '" & Replace(vFldDao(Rs("Nombre")), Chr(39), "") & "'"
                    Q1 = Q1 & " ,Direccion = '" & vFldDao(Rs("Direccion")) & "'"
                    Q1 = Q1 & " ,Region = " & vFldDao(Rs("Region"))
                    Q1 = Q1 & " ,Comuna = " & vFldDao(Rs("Comuna"))
                    Q1 = Q1 & " ,Ciudad = '" & vFldDao(Rs("Ciudad")) & "'"
                    Q1 = Q1 & " ,Telefonos = '" & vFldDao(Rs("Telefonos")) & "'"
                    Q1 = Q1 & " ,Fax = '" & vFldDao(Rs("Fax")) & "'"
                    Q1 = Q1 & " ,ActEcon = " & vFldDao(Rs("ActEcon"))
                    Q1 = Q1 & " ,CodActEcon = '" & vFldDao(Rs("CodActEcon")) & "'"
                    Q1 = Q1 & " ,DomPostal = '" & vFldDao(Rs("DomPostal")) & "'"
                    Q1 = Q1 & " ,ComPostal = " & vFldDao(Rs("ComPostal"))
                    Q1 = Q1 & " ,Email = '" & vFldDao(Rs("Email")) & "'"
                    Q1 = Q1 & " ,Web = '" & vFldDao(Rs("Web")) & "'"
                    Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                    Q1 = Q1 & " ,Obs = '" & vFldDao(Rs("Obs")) & "'"
                    Q1 = Q1 & " ,Clasif0 = " & vFldDao(Rs("Clasif0"))
                    Q1 = Q1 & " ,Clasif1 = " & vFldDao(Rs("Clasif1"))
                    Q1 = Q1 & " ,Clasif2 = " & vFldDao(Rs("Clasif2"))
                    Q1 = Q1 & " ,Clasif3 = " & vFldDao(Rs("Clasif3"))
                    Q1 = Q1 & " ,Clasif4 = " & vFldDao(Rs("Clasif4"))
                    Q1 = Q1 & " ,Clasif5 = " & vFldDao(Rs("Clasif5"))
                    Q1 = Q1 & " ,Giro = '" & vFldDao(Rs("Giro")) & "'"
                    Q1 = Q1 & " ,NotValidRut = " & vFldDao(Rs("NotValidRut"))
                    Q1 = Q1 & " ,EsSupermercado = " & vFldDao(Rs("EsSupermercado"))
                    Q1 = Q1 & " ,EntRelacionada = " & vFldDao(Rs("EntRelacionada"))
                    Q1 = Q1 & " ,CodCtaAfecto = '" & vFldDao(Rs("CodCtaAfecto")) & "'"
                    Q1 = Q1 & " ,CodCtaExento = '" & vFldDao(Rs("CodCtaExento")) & "'"
                    Q1 = Q1 & " ,CodCtaTotal = '" & vFldDao(Rs("CodCtaTotal")) & "'"
                    Q1 = Q1 & " ,PropIVA = " & vFldDao(Rs("PropIVA"))
                    Q1 = Q1 & " ,CodCCostoAfecto = '" & vFldDao(Rs("CodCCostoAfecto")) & "'"
                    Q1 = Q1 & " ,CodAreaNegAfecto = '" & vFldDao(Rs("CodAreaNegAfecto")) & "'"
                    Q1 = Q1 & " ,CodCCostoExento = '" & vFldDao(Rs("CodCCostoExento")) & "'"
                    Q1 = Q1 & " ,CodAreaNegExento = '" & vFldDao(Rs("CodAreaNegExento")) & "'"
                    Q1 = Q1 & " ,CodCCostoTotal = '" & vFldDao(Rs("CodCCostoTotal")) & "'"
                    Q1 = Q1 & " ,CodAreaNegTotal = '" & vFldDao(Rs("CodAreaNegTotal")) & "'"
                    Q1 = Q1 & " ,CodCtaAfectoVta = '" & vFldDao(Rs("CodCtaAfectoVta")) & "'"
                    Q1 = Q1 & " ,CodCtaExentoVta = '" & vFldDao(Rs("CodCtaExentoVta")) & "'"
                    Q1 = Q1 & " ,CodCtaTotalVta = '" & vFldDao(Rs("CodCtaTotalVta")) & "'"
                    Q1 = Q1 & " ,CodCCostoAfectoVta = '" & vFldDao(Rs("CodCCostoAfectoVta")) & "'"
                    Q1 = Q1 & " ,CodAreaNegAfectoVta = '" & vFldDao(Rs("CodAreaNegAfectoVta")) & "'"
                    Q1 = Q1 & " ,CodCCostoExentoVta = '" & vFldDao(Rs("CodCCostoExentoVta")) & "'"
                    Q1 = Q1 & " ,CodAreaNegExentoVta = '" & vFldDao(Rs("CodAreaNegExentoVta")) & "'"
                    Q1 = Q1 & " ,CodCCostoTotalVta = '" & vFldDao(Rs("CodCCostoTotalVta")) & "'"
                    Q1 = Q1 & " ,CodAreaNegTotalVta = '" & vFldDao(Rs("CodAreaNegTotalVta")) & "'"
                    Q1 = Q1 & " ,EsDelGiro = " & vFldDao(Rs("EsDelGiro"))
                    Q1 = Q1 & " ,FranqTribEnt = " & vFldDao(Rs("FranqTribEnt"))
                    Q1 = Q1 & " ,Ret3Porc = " & vFldDao(Rs("Ret3Porc"))
                    Q1 = Q1 & " ,FDesde3Porc = " & vFldDao(Rs("FDesde3Porc"))
                    Q1 = Q1 & " ,FHasta3Porc = " & vFldDao(Rs("FHasta3Porc"))
                    Q1 = Q1 & " WHERE Rut = '" & Replace(Replace(vFldDao(Rs("Rut")), "-", ""), ".", "") & "'"
                    Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa
                    Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdEntidad"))
                    Call ExecSQL(DBSql, Q1)
                    
                End If
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
   If Err Then
      Q1 = ""
   End If


End Sub

Public Sub TrasCuentas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Cuentas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Cuentas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)
    
    
   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,idCuenta"
   Q1 = Q1 & " ,idPadre"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,CodFECU"
   Q1 = Q1 & " ,Nivel"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,MarcaApertura"
   Q1 = Q1 & " ,TipoCapPropio"
   Q1 = Q1 & " ,CodF22"
   Q1 = Q1 & " ,Atrib1"
   Q1 = Q1 & " ,Atrib2"
   Q1 = Q1 & " ,Atrib3"
   Q1 = Q1 & " ,Atrib4"
   Q1 = Q1 & " ,Atrib5"
   Q1 = Q1 & " ,Atrib6"
   Q1 = Q1 & " ,Atrib7"
   Q1 = Q1 & " ,Atrib8"
   Q1 = Q1 & " ,Atrib9"
   Q1 = Q1 & " ,Atrib10"
   Q1 = Q1 & " ,CodF29"
   Q1 = Q1 & " ,CorrelativoCheque"
   Q1 = Q1 & " ,CodIFRS_EstRes"
   Q1 = Q1 & " ,CodIFRS_EstFin"
   Q1 = Q1 & " ,DebeTrib"
   Q1 = Q1 & " ,HaberTrib"
   Q1 = Q1 & " ,CodIFRS"
   Q1 = Q1 & " ,CodF22_14Ter"
   Q1 = Q1 & " ,TipoPartida"
   Q1 = Q1 & " ,CodCtaPlanSII"
   Q1 = Q1 & " ,IdCuentaOld"
   Q1 = Q1 & " ,IdPadreOld"
   'Q1 = Q1 & " ,IdTras"
   Q1 = Q1 & " FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Cuentas"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCuenta"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Cuentas ON "
                Q1 = " INSERT INTO Cuentas"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,idPadre"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,CodFECU"
                Q1 = Q1 & " ,Nivel"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Clasificacion"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,MarcaApertura"
                Q1 = Q1 & " ,TipoCapPropio"
                Q1 = Q1 & " ,CodF22"
                Q1 = Q1 & " ,Atrib1"
                Q1 = Q1 & " ,Atrib2"
                Q1 = Q1 & " ,Atrib3"
                Q1 = Q1 & " ,Atrib4"
                Q1 = Q1 & " ,Atrib5"
                Q1 = Q1 & " ,Atrib6"
                Q1 = Q1 & " ,Atrib7"
                Q1 = Q1 & " ,Atrib8"
                Q1 = Q1 & " ,Atrib9"
                Q1 = Q1 & " ,Atrib10"
                Q1 = Q1 & " ,CodF29"
                Q1 = Q1 & " ,CorrelativoCheque"
                Q1 = Q1 & " ,CodIFRS_EstRes"
                Q1 = Q1 & " ,CodIFRS_EstFin"
                Q1 = Q1 & " ,DebeTrib"
                Q1 = Q1 & " ,HaberTrib"
                Q1 = Q1 & " ,CodIFRS"
                Q1 = Q1 & " ,CodF22_14Ter"
                Q1 = Q1 & " ,TipoPartida"
                Q1 = Q1 & " ,CodCtaPlanSII"
                Q1 = Q1 & " ,IdCuentaOld"
                Q1 = Q1 & " ,IdPadreOld"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ," & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29"))
                Q1 = Q1 & " ," & vFldDao(Rs("CorrelativoCheque"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("DebeTrib"))
                Q1 = Q1 & " ," & vFldDao(Rs("HaberTrib"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("CodF22_14Ter"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdPadreOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Cuentas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Cuentas"
                Q1 = Q1 & " SET idPadre = " & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,CodFECU = '" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,MarcaApertura = " & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ,TipoCapPropio = " & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ,CodF22 = " & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ,Atrib1 = " & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ,Atrib2 = " & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ,Atrib3 = " & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ,Atrib4 = " & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ,Atrib5 = " & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ,Atrib6 = " & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ,Atrib7 = " & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ,Atrib8 = " & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ,Atrib9 = " & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ,Atrib10 = " & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,CodF29 = " & vFldDao(Rs("CodF29"))
                Q1 = Q1 & " ,CorrelativoCheque = " & vFldDao(Rs("CorrelativoCheque"))
                Q1 = Q1 & " ,CodIFRS_EstRes = '" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,CodIFRS_EstFin = '" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ,DebeTrib = " & vFldDao(Rs("DebeTrib"))
                Q1 = Q1 & " ,HaberTrib = " & vFldDao(Rs("HaberTrib"))
                Q1 = Q1 & " ,CodIFRS = '" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ,CodF22_14Ter = " & vFldDao(Rs("CodF22_14Ter"))
                Q1 = Q1 & " ,TipoPartida = " & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,CodCtaPlanSII = '" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                Q1 = Q1 & " ,IdCuentaOld = " & vFldDao(Rs("IdCuentaOld"))
                Q1 = Q1 & " ,IdPadreOld = " & vFldDao(Rs("IdPadreOld"))
                Q1 = Q1 & " ,IdTras = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCuenta"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    'ACTUALIZA LAS CUENTAS PADRES CON EL NUEVO ID DE CUENTAS
    Q1 = " UPDATE CUE"
    Q1 = Q1 & " Set CUE.IdPadre = CU.IdCuenta"
    Q1 = Q1 & " FROM Cuentas AS CU INNER JOIN Cuentas AS CUE ON CU.IDTRAS = CUE.idPadre AND CU.IDEMPRESA = CUE.IDEMPRESA AND CU.ANO = CUE.ANO"
    Q1 = Q1 & " WHERE CU.IdEmpresa = " & IdEmpresa & " And CU.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasUsuarios(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM Usuarios "
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Usuarios"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT IdUsuario"
   Q1 = Q1 & " ,Usuario"
   Q1 = Q1 & " ,Clave"
   Q1 = Q1 & " ,NombreLargo"
   Q1 = Q1 & " ,PrivAdm"
   Q1 = Q1 & " ,Activo"
   Q1 = Q1 & " ,HabilitadoHasta"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Usuarios) as Cant"
   Q1 = Q1 & " FROM Usuarios "
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Usuarios"
            Q1 = Q1 & " WHERE IdUsuario = " & vFldDao(Rs("IdUsuario"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                Q1 = " SET IDENTITY_INSERT Usuarios ON "
                Q1 = Q1 & " INSERT INTO Usuarios"
                Q1 = Q1 & " (IdUsuario"
                Q1 = Q1 & " ,Usuario"
                Q1 = Q1 & " ,Clave"
                Q1 = Q1 & " ,NombreLargo"
                Q1 = Q1 & " ,PrivAdm"
                Q1 = Q1 & " ,Activo"
                Q1 = Q1 & " ,HabilitadoHasta)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Usuario")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Clave"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombreLargo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("PrivAdm"))
                Q1 = Q1 & " ," & vFldDao(Rs("Activo"))
                Q1 = Q1 & " ," & vFldDao(Rs("HabilitadoHasta")) & ")"
                Q1 = Q1 & " SET IDENTITY_INSERT Usuarios OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Usuarios"
                Q1 = Q1 & " SET Usuario = '" & vFldDao(Rs("Usuario")) & "'"
                Q1 = Q1 & " ,Clave = " & vFldDao(Rs("Clave"))
                Q1 = Q1 & " ,NombreLargo = '" & vFldDao(Rs("NombreLargo")) & "'"
                Q1 = Q1 & " ,PrivAdm = " & vFldDao(Rs("PrivAdm"))
                Q1 = Q1 & " ,Activo = " & vFldDao(Rs("Activo"))
                Q1 = Q1 & " ,HabilitadoHasta = " & vFldDao(Rs("HabilitadoHasta"))
                Q1 = Q1 & " WHERE IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasUsuarioEmpresa(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From UsuarioEmpresa"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)
    

   Q1 = "SELECT idUsuario"
   Q1 = Q1 & " ,idEmpresa"
   Q1 = Q1 & " ,idPerfil"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM UsuarioEmpresa) as Cant"
   Q1 = Q1 & " From UsuarioEmpresa"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From UsuarioEmpresa"
            Q1 = Q1 & " WHERE idUsuario = " & vFldDao(Rs("idUsuario"))
            Q1 = Q1 & " AND idEmpresa = " & gEmpresa.id
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT UsuarioEmpresa ON "
                Q1 = " INSERT INTO UsuarioEmpresa"
                Q1 = Q1 & " (idUsuario"
                Q1 = Q1 & " ,idEmpresa"
                Q1 = Q1 & " ,idPerfil)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idUsuario"))
                Q1 = Q1 & " ," & gEmpresa.id
                Q1 = Q1 & " ," & vFldDao(Rs("idPerfil")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT UsuarioEmpresa OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE UsuarioEmpresa"
                Q1 = Q1 & " SET idPerfil = " & vFldDao(Rs("idPerfil"))
                Q1 = Q1 & " WHERE idUsuario = " & vFldDao(Rs("idUsuario"))
                Q1 = Q1 & " AND idEmpresa = " & gEmpresa.id
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasTimbraje(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Timbraje"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idEmpresa"
   Q1 = Q1 & " ,UltImpreso"
   Q1 = Q1 & " ,FUltImpreso"
   Q1 = Q1 & " ,UltTimbrado"
   Q1 = Q1 & " ,FUltTimbrado"
   Q1 = Q1 & " ,UltUsado"
   Q1 = Q1 & " ,FUltUsado"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Timbraje) as Cant"
   Q1 = Q1 & " From Timbraje"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Timbraje"
            Q1 = Q1 & " WHERE idEmpresa = " & gEmpresa.id
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                'Q1 = " SET IDENTITY_INSERT UsuarioEmpresa ON "
                Q1 = " INSERT INTO Timbraje"
                Q1 = Q1 & " (idEmpresa"
                Q1 = Q1 & " ,UltImpreso"
                Q1 = Q1 & " ,FUltImpreso"
                Q1 = Q1 & " ,UltTimbrado"
                Q1 = Q1 & " ,FUltTimbrado"
                Q1 = Q1 & " ,UltUsado"
                Q1 = Q1 & " ,FUltUsado)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & gEmpresa.id
                Q1 = Q1 & " ," & vFldDao(Rs("UltImpreso"))
                Q1 = Q1 & " ," & vFldDao(Rs("FUltImpreso"))
                Q1 = Q1 & " ," & vFldDao(Rs("UltTimbrado"))
                Q1 = Q1 & " ," & vFldDao(Rs("FUltTimbrado"))
                Q1 = Q1 & " ," & vFldDao(Rs("UltUsado"))
                Q1 = Q1 & " ," & vFldDao(Rs("FUltUsado")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT UsuarioEmpresa OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Timbraje"
                Q1 = Q1 & " SET UltImpreso = " & vFldDao(Rs("UltImpreso"))
                Q1 = Q1 & " ,FUltImpreso = " & vFldDao(Rs("FUltImpreso"))
                Q1 = Q1 & " ,UltTimbrado = " & vFldDao(Rs("UltTimbrado"))
                Q1 = Q1 & " ,FUltTimbrado = " & vFldDao(Rs("FUltTimbrado"))
                Q1 = Q1 & " ,UltUsado = " & vFldDao(Rs("UltUsado"))
                Q1 = Q1 & " ,FUltUsado = " & vFldDao(Rs("FUltUsado"))
                Q1 = Q1 & " WHERE idEmpresa = " & gEmpresa.id
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasTipoDocs(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From TipoDocs"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)
    
'    Q1 = Q1 & " DELETE FROM TipoDocs "
'    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT Id"
   Q1 = Q1 & " ,TipoLib"
   Q1 = Q1 & " ,TipoDoc"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Diminutivo"
   Q1 = Q1 & " ,Atributo"
   Q1 = Q1 & " ,TipoDocFijo"
   Q1 = Q1 & " ,CodF29Count"
   Q1 = Q1 & " ,CodF29Neto"
   Q1 = Q1 & " ,CodF29IVA"
   Q1 = Q1 & " ,CodF29IVADTE"
   Q1 = Q1 & " ,CodF29AFCount"
   Q1 = Q1 & " ,CodF29AFIVA"
   Q1 = Q1 & " ,CodF29RetHon"
   Q1 = Q1 & " ,CodF29RetDieta"
   Q1 = Q1 & " ,CodF29IVARet3ro"
   Q1 = Q1 & " ,TieneAfecto"
   Q1 = Q1 & " ,TieneExento"
   Q1 = Q1 & " ,ExigeRUT"
   Q1 = Q1 & " ,EsRebaja"
   Q1 = Q1 & " ,DocImpExp"
   Q1 = Q1 & " ,DocBoletas"
   Q1 = Q1 & " ,CodF29ExCount"
   Q1 = Q1 & " ,CodF29Exento"
   Q1 = Q1 & " ,CodF29CountNoGiro"
   Q1 = Q1 & " ,CodF29NetoNoGiro"
   Q1 = Q1 & " ,CodF29IVANoGiro"
   Q1 = Q1 & " ,CodF29ExCountNoGiro"
   Q1 = Q1 & " ,CodF29ExentoNoGiro"
   Q1 = Q1 & " ,CodF29CountRetParcial"
   Q1 = Q1 & " ,CodF29NetoRetParcial"
   Q1 = Q1 & " ,CodF29DifIVARetParcial"
   Q1 = Q1 & " ,CodF29CountDTE"
   Q1 = Q1 & " ,CodF29NetoDTE"
   Q1 = Q1 & " ,CodF29IVAIrrecDTE"
   Q1 = Q1 & " ,CodF29CountIVAIrrec"
   Q1 = Q1 & " ,CodF29NetoIVAIrrec"
   Q1 = Q1 & " ,CodDocSII"
   Q1 = Q1 & " ,CodDocDTESII"
   Q1 = Q1 & " ,AceptaPropIVA"
   Q1 = Q1 & " ,CodF29CountSuper"
   Q1 = Q1 & " ,CodF29IVASuper"
   Q1 = Q1 & " ,IngresarTotal"
   Q1 = Q1 & " ,TieneNumDocHasta"
   Q1 = Q1 & " ,TieneCantBoletas"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM TipoDocs) as Cant"
   Q1 = Q1 & " From TipoDocs"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
    If Rs.EOF = False Then
    
       If CantSql < vFldDao(Rs("Cant")) Then
       
            Q1 = Q1 & " DELETE FROM TipoDocs "
            Call ExecSQL(DBSql, Q1)
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From TipoDocs"
            Q1 = Q1 & " WHERE Id = " & vFldDao(Rs("Id"))
            Q1 = Q1 & " AND TipoLib = " & vFldDao(Rs("TipoLib"))
            Q1 = Q1 & " AND TipoDoc = " & vFldDao(Rs("TipoDoc"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT TipoDocs ON "
                Q1 = Q1 & " INSERT INTO TipoDocs"
                Q1 = Q1 & " (Id"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,TipoDoc"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Diminutivo"
                Q1 = Q1 & " ,Atributo"
                Q1 = Q1 & " ,TipoDocFijo"
                Q1 = Q1 & " ,CodF29Count"
                Q1 = Q1 & " ,CodF29Neto"
                Q1 = Q1 & " ,CodF29IVA"
                Q1 = Q1 & " ,CodF29IVADTE"
                Q1 = Q1 & " ,CodF29AFCount"
                Q1 = Q1 & " ,CodF29AFIVA"
                Q1 = Q1 & " ,CodF29RetHon"
                Q1 = Q1 & " ,CodF29RetDieta"
                Q1 = Q1 & " ,CodF29IVARet3ro"
                Q1 = Q1 & " ,TieneAfecto"
                Q1 = Q1 & " ,TieneExento"
                Q1 = Q1 & " ,ExigeRUT"
                Q1 = Q1 & " ,EsRebaja"
                Q1 = Q1 & " ,DocImpExp"
                Q1 = Q1 & " ,DocBoletas"
                Q1 = Q1 & " ,CodF29ExCount"
                Q1 = Q1 & " ,CodF29Exento"
                Q1 = Q1 & " ,CodF29CountNoGiro"
                Q1 = Q1 & " ,CodF29NetoNoGiro"
                Q1 = Q1 & " ,CodF29IVANoGiro"
                Q1 = Q1 & " ,CodF29ExCountNoGiro"
                Q1 = Q1 & " ,CodF29ExentoNoGiro"
                Q1 = Q1 & " ,CodF29CountRetParcial"
                Q1 = Q1 & " ,CodF29NetoRetParcial"
                Q1 = Q1 & " ,CodF29DifIVARetParcial"
                Q1 = Q1 & " ,CodF29CountDTE"
                Q1 = Q1 & " ,CodF29NetoDTE"
                Q1 = Q1 & " ,CodF29IVAIrrecDTE"
                Q1 = Q1 & " ,CodF29CountIVAIrrec"
                Q1 = Q1 & " ,CodF29NetoIVAIrrec"
                Q1 = Q1 & " ,CodDocSII"
                Q1 = Q1 & " ,CodDocDTESII"
                Q1 = Q1 & " ,AceptaPropIVA"
                Q1 = Q1 & " ,CodF29CountSuper"
                Q1 = Q1 & " ,CodF29IVASuper"
                Q1 = Q1 & " ,IngresarTotal"
                Q1 = Q1 & " ,TieneNumDocHasta"
                Q1 = Q1 & " ,TieneCantBoletas)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("Id"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Atributo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDocFijo"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29Count"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29Neto"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVADTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29AFCount"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29AFIVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29RetHon"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29RetDieta"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVARet3ro"))
                Q1 = Q1 & " ," & vFldDao(Rs("TieneAfecto"))
                Q1 = Q1 & " ," & vFldDao(Rs("TieneExento"))
                Q1 = Q1 & " ," & vFldDao(Rs("ExigeRUT"))
                Q1 = Q1 & " ," & vFldDao(Rs("EsRebaja"))
                Q1 = Q1 & " ," & vFldDao(Rs("DocImpExp"))
                Q1 = Q1 & " ," & vFldDao(Rs("DocBoletas"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29ExCount"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29Exento"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29CountNoGiro"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29NetoNoGiro"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVANoGiro"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29ExCountNoGiro"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29ExentoNoGiro"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29CountRetParcial"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29NetoRetParcial"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29DifIVARetParcial"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29CountDTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29NetoDTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVAIrrecDTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29CountIVAIrrec"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29NetoIVAIrrec"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodDocSII")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodDocDTESII")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("AceptaPropIVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29CountSuper"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29IVASuper"))
                Q1 = Q1 & " ," & vFldDao(Rs("IngresarTotal"))
                Q1 = Q1 & " ," & vFldDao(Rs("TieneNumDocHasta"))
                Q1 = Q1 & " ," & vFldDao(Rs("TieneCantBoletas")) & ")"
                Q1 = Q1 & " SET IDENTITY_INSERT TipoDocs OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE TipoDocs"
                Q1 = Q1 & " SET TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,TipoDoc = " & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Diminutivo = '" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,Atributo = '" & vFldDao(Rs("Atributo")) & "'"
                Q1 = Q1 & " ,TipoDocFijo = " & vFldDao(Rs("TipoDocFijo"))
                Q1 = Q1 & " ,CodF29Count = " & vFldDao(Rs("CodF29Count"))
                Q1 = Q1 & " ,CodF29Neto = " & vFldDao(Rs("CodF29Neto"))
                Q1 = Q1 & " ,CodF29IVA = " & vFldDao(Rs("CodF29IVA"))
                Q1 = Q1 & " ,CodF29IVADTE = " & vFldDao(Rs("CodF29IVADTE"))
                Q1 = Q1 & " ,CodF29AFCount = " & vFldDao(Rs("CodF29AFCount"))
                Q1 = Q1 & " ,CodF29AFIVA = " & vFldDao(Rs("CodF29AFIVA"))
                Q1 = Q1 & " ,CodF29RetHon = " & vFldDao(Rs("CodF29RetHon"))
                Q1 = Q1 & " ,CodF29RetDieta = " & vFldDao(Rs("CodF29RetDieta"))
                Q1 = Q1 & " ,CodF29IVARet3ro = " & vFldDao(Rs("CodF29IVARet3ro"))
                Q1 = Q1 & " ,TieneAfecto = " & vFldDao(Rs("TieneAfecto"))
                Q1 = Q1 & " ,TieneExento = " & vFldDao(Rs("TieneExento"))
                Q1 = Q1 & " ,ExigeRUT = " & vFldDao(Rs("ExigeRUT"))
                Q1 = Q1 & " ,EsRebaja = " & vFldDao(Rs("EsRebaja"))
                Q1 = Q1 & " ,DocImpExp = " & vFldDao(Rs("DocImpExp"))
                Q1 = Q1 & " ,DocBoletas = " & vFldDao(Rs("DocBoletas"))
                Q1 = Q1 & " ,CodF29ExCount = " & vFldDao(Rs("CodF29ExCount"))
                Q1 = Q1 & " ,CodF29Exento = " & vFldDao(Rs("CodF29Exento"))
                Q1 = Q1 & " ,CodF29CountNoGiro = " & vFldDao(Rs("CodF29CountNoGiro"))
                Q1 = Q1 & " ,CodF29NetoNoGiro = " & vFldDao(Rs("CodF29NetoNoGiro"))
                Q1 = Q1 & " ,CodF29IVANoGiro = " & vFldDao(Rs("CodF29IVANoGiro"))
                Q1 = Q1 & " ,CodF29ExCountNoGiro = " & vFldDao(Rs("CodF29ExCountNoGiro"))
                Q1 = Q1 & " ,CodF29ExentoNoGiro = " & vFldDao(Rs("CodF29ExentoNoGiro"))
                Q1 = Q1 & " ,CodF29CountRetParcial = " & vFldDao(Rs("CodF29CountRetParcial"))
                Q1 = Q1 & " ,CodF29NetoRetParcial = " & vFldDao(Rs("CodF29NetoRetParcial"))
                Q1 = Q1 & " ,CodF29DifIVARetParcial = " & vFldDao(Rs("CodF29DifIVARetParcial"))
                Q1 = Q1 & " ,CodF29CountDTE = " & vFldDao(Rs("CodF29CountDTE"))
                Q1 = Q1 & " ,CodF29NetoDTE = " & vFldDao(Rs("CodF29NetoDTE"))
                Q1 = Q1 & " ,CodF29IVAIrrecDTE = " & vFldDao(Rs("CodF29IVAIrrecDTE"))
                Q1 = Q1 & " ,CodF29CountIVAIrrec = " & vFldDao(Rs("CodF29CountIVAIrrec"))
                Q1 = Q1 & " ,CodF29NetoIVAIrrec = " & vFldDao(Rs("CodF29NetoIVAIrrec"))
                Q1 = Q1 & " ,CodDocSII = '" & vFldDao(Rs("CodDocSII")) & "'"
                Q1 = Q1 & " ,CodDocDTESII = '" & vFldDao(Rs("CodDocDTESII")) & "'"
                Q1 = Q1 & " ,AceptaPropIVA = " & vFldDao(Rs("AceptaPropIVA"))
                Q1 = Q1 & " ,CodF29CountSuper = " & vFldDao(Rs("CodF29CountSuper"))
                Q1 = Q1 & " ,CodF29IVASuper = " & vFldDao(Rs("CodF29IVASuper"))
                Q1 = Q1 & " ,IngresarTotal = " & vFldDao(Rs("IngresarTotal"))
                Q1 = Q1 & " ,TieneNumDocHasta = " & vFldDao(Rs("TieneNumDocHasta"))
                Q1 = Q1 & " ,TieneCantBoletas = " & vFldDao(Rs("TieneCantBoletas"))
                Q1 = Q1 & " WHERE Id = " & vFldDao(Rs("Id"))
                Q1 = Q1 & " AND TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " AND TipoDoc = " & vFldDao(Rs("TipoDoc"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasSucursales(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Sucursales' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Sucursales ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdSucursal"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,Vigente"
   Q1 = Q1 & " From Sucursales"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Sucursales"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdSucursal"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Sucursales ON "
                Q1 = " INSERT INTO Sucursales"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,Vigente"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdSucursal")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Sucursales OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Sucursales"
                Q1 = Q1 & " SET Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,Vigente = " & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdSucursal"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasRegiones(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
'    Q1 = "DELETE FROM Regiones "
'    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Regiones"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT Id"
   Q1 = Q1 & " ,CODIGO"
   Q1 = Q1 & " ,COMUNA"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Regiones) as Cant"
   Q1 = Q1 & " From Regiones"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Regiones"
            Q1 = Q1 & " WHERE Id = " & vFldDao(Rs("Id"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                Q1 = " SET IDENTITY_INSERT Regiones ON "
                Q1 = Q1 & " INSERT INTO Regiones"
                Q1 = Q1 & " (Id"
                Q1 = Q1 & " ,CODIGO"
                Q1 = Q1 & " ,COMUNA)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("Id"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CODIGO")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("COMUNA")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT Regiones OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Regiones"
                Q1 = Q1 & " SET CODIGO = '" & vFldDao(Rs("CODIGO")) & "'"
                Q1 = Q1 & " ,COMUNA = '" & vFldDao(Rs("COMUNA")) & "'"
                Q1 = Q1 & " WHERE Id = " & vFldDao(Rs("Id"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasRazonesFin(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM RazonesFin"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From RazonesFin"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT IdRazon"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,RazonFija"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,UnidadRes"
   Q1 = Q1 & " ,TxtNumerador"
   Q1 = Q1 & " ,TxtDenominador"
   Q1 = Q1 & " ,Operador"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM RazonesFin) as Cant"
   Q1 = Q1 & " From RazonesFin"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From RazonesFin"
            Q1 = Q1 & " WHERE IdRazon = " & vFldDao(Rs("IdRazon"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT RazonesFin ON "
                Q1 = Q1 & " INSERT INTO RazonesFin"
                Q1 = Q1 & " (IdRazon"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,RazonFija"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,UnidadRes"
                Q1 = Q1 & " ,TxtNumerador"
                Q1 = Q1 & " ,TxtDenominador"
                Q1 = Q1 & " ,Operador"
                Q1 = Q1 & " ,Glosa)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdRazon"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ," & vFldDao(Rs("RazonFija"))
                Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Nombre")), Chr(39), "¤") & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("UnidadRes")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("TxtNumerador")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("TxtDenominador")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Operador")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Glosa")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT RazonesFin OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE RazonesFin"
                Q1 = Q1 & " SET Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,RazonFija = " & vFldDao(Rs("RazonFija"))
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,UnidadRes = '" & vFldDao(Rs("UnidadRes")) & "'"
                Q1 = Q1 & " ,TxtNumerador = '" & vFldDao(Rs("TxtNumerador")) & "'"
                Q1 = Q1 & " ,TxtDenominador = '" & vFldDao(Rs("TxtDenominador")) & "'"
                Q1 = Q1 & " ,Operador = '" & vFldDao(Rs("Operador")) & "'"
                Q1 = Q1 & " ,Glosa = '" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " WHERE IdRazon = " & vFldDao(Rs("IdRazon"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasPropIVA_TotMensual(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   

   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Mes"
   Q1 = Q1 & " ,TotalAfecto"
   Q1 = Q1 & " ,TotalExento"
   Q1 = Q1 & " From PropIVA_TotMensual"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From PropIVA_TotMensual"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT PropIVA_TotMensual ON "
                Q1 = " INSERT INTO PropIVA_TotMensual"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Mes"
                Q1 = Q1 & " ,TotalAfecto"
                Q1 = Q1 & " ,TotalExento)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("Mes"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("TotalAfecto"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("TotalExento")))) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT PropIVA_TotMensual OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE PropIVA_TotMensual"
                Q1 = Q1 & " SET TotalAfecto = " & str(vFmt(vFldDao(Rs("TotalAfecto"))))
                Q1 = Q1 & " ,TotalExento = " & str(vFmt(vFldDao(Rs("TotalExento"))))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasPlanIntermedio(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM PlanIntermedio"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From PlanIntermedio"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idCuenta"
   Q1 = Q1 & " ,idPadre"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,CodFECU"
   Q1 = Q1 & " ,Nivel"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,MarcaApertura"
   Q1 = Q1 & " ,TipoCapPropio"
   Q1 = Q1 & " ,CodF22"
   Q1 = Q1 & " ,Atrib1"
   Q1 = Q1 & " ,Atrib2"
   Q1 = Q1 & " ,Atrib3"
   Q1 = Q1 & " ,Atrib4"
   Q1 = Q1 & " ,Atrib5"
   Q1 = Q1 & " ,Atrib6"
   Q1 = Q1 & " ,Atrib7"
   Q1 = Q1 & " ,Atrib8"
   Q1 = Q1 & " ,Atrib9"
   Q1 = Q1 & " ,Atrib10"
   Q1 = Q1 & " ,CodIFRS_EstRes"
   Q1 = Q1 & " ,CodIFRS_EstFin"
   Q1 = Q1 & " ,CodIFRS"
   Q1 = Q1 & " ,TipoPartida"
   Q1 = Q1 & " ,CodCtaPlanSII"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM PlanIntermedio) as Cant"
   Q1 = Q1 & " From PlanIntermedio"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From PlanIntermedio"
            Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT PlanIntermedio ON "
                Q1 = Q1 & " INSERT INTO PlanIntermedio"
                Q1 = Q1 & " (idCuenta"
                Q1 = Q1 & " ,idPadre"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,CodFECU"
                Q1 = Q1 & " ,Nivel"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Clasificacion"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,MarcaApertura"
                Q1 = Q1 & " ,TipoCapPropio"
                Q1 = Q1 & " ,CodF22"
                Q1 = Q1 & " ,Atrib1"
                Q1 = Q1 & " ,Atrib2"
                Q1 = Q1 & " ,Atrib3"
                Q1 = Q1 & " ,Atrib4"
                Q1 = Q1 & " ,Atrib5"
                Q1 = Q1 & " ,Atrib6"
                Q1 = Q1 & " ,Atrib7"
                Q1 = Q1 & " ,Atrib8"
                Q1 = Q1 & " ,Atrib9"
                Q1 = Q1 & " ,Atrib10"
                Q1 = Q1 & " ,CodIFRS_EstRes"
                Q1 = Q1 & " ,CodIFRS_EstFin"
                Q1 = Q1 & " ,CodIFRS"
                Q1 = Q1 & " ,TipoPartida"
                Q1 = Q1 & " ,CodCtaPlanSII)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ," & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaPlanSII")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT PlanIntermedio OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE PlanIntermedio"
                Q1 = Q1 & " SET idPadre = " & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,CodFECU = '" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,MarcaApertura = " & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ,TipoCapPropio = " & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ,CodF22 = " & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ,Atrib1 = " & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ,Atrib2 = " & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ,Atrib3 = " & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ,Atrib4 = " & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ,Atrib5 = " & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ,Atrib6 = " & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ,Atrib7 = " & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ,Atrib8 = " & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ,Atrib9 = " & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ,Atrib10 = " & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,CodIFRS_EstRes = '" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,CodIFRS_EstFin = '" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ,CodIFRS = '" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ,TipoPartida = " & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,CodCtaPlanSII = '" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasPerfiles(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Perfiles"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT IdPerfil"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Privilegios"
   Q1 = Q1 & " ,IdApp"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Perfiles) as Cant"
   Q1 = Q1 & " From Perfiles"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Perfiles"
            Q1 = Q1 & " WHERE IdPerfil = " & vFldDao(Rs("IdPerfil"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Perfiles ON "
                Q1 = " INSERT INTO Perfiles"
                Q1 = Q1 & " (IdPerfil"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Privilegios"
                Q1 = Q1 & " ,IdApp)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdPerfil"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Privilegios"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdApp")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Perfiles OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Perfiles"
                Q1 = Q1 & " SET Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Privilegios = " & vFldDao(Rs("Privilegios"))
                Q1 = Q1 & " ,IdApp = " & vFldDao(Rs("IdApp"))
                Q1 = Q1 & " WHERE IdPerfil = " & vFldDao(Rs("IdPerfil"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)

End Sub

Public Sub TrasPercepciones(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Percepciones' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Percepciones ADD IdTras INT NULL; "
    Q1 = Q1 & "END "

    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IDPerc"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,Orden"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,NumCertificado"
   Q1 = Q1 & " ,RutEmpresa"
   Q1 = Q1 & " ,Regimen"
   Q1 = Q1 & " ,Contabilizacion"
   Q1 = Q1 & " ,TasaTef"
   Q1 = Q1 & " ,TasaTex"
   Q1 = Q1 & " ,Percepciones"
   Q1 = Q1 & " From Percepciones"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Percepciones"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IDPerc"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                
                Call CloseRs(Rs1)
                Q1 = "SELECT MAX(IDPerc) + 1 AS  IDPerc"
                Q1 = Q1 & " From Percepciones"
                Set Rs1 = OpenRs(DBSql, Q1)
            
                'Q1 = " SET IDENTITY_INSERT Percepciones ON "
                Q1 = " INSERT INTO Percepciones"
                Q1 = Q1 & " (IDPerc"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,Orden"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,NumCertificado"
                Q1 = Q1 & " ,RutEmpresa"
                Q1 = Q1 & " ,Regimen"
                Q1 = Q1 & " ,Contabilizacion"
                Q1 = Q1 & " ,TasaTef"
                Q1 = Q1 & " ,TasaTex"
                Q1 = Q1 & " ,Percepciones"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IIf(vFld(Rs1("IDPerc")) = 0, 1, vFld(Rs1("IDPerc")))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumCertificado"))
                Q1 = Q1 & " ," & vFldDao(Rs("RutEmpresa"))
                Q1 = Q1 & " ," & vFldDao(Rs("Regimen"))
                Q1 = Q1 & " ," & vFldDao(Rs("Contabilizacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("TasaTef"))
                Q1 = Q1 & " ," & vFldDao(Rs("TasaTex"))
                Q1 = Q1 & " ," & vFldDao(Rs("Percepciones"))
                Q1 = Q1 & " ," & vFldDao(Rs("IDPerc")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Percepciones OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Percepciones"
                Q1 = Q1 & " SET IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,Orden = " & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,NumCertificado = " & vFldDao(Rs("NumCertificado"))
                Q1 = Q1 & " ,RutEmpresa = " & vFldDao(Rs("RutEmpresa"))
                Q1 = Q1 & " ,Regimen = " & vFldDao(Rs("Regimen"))
                Q1 = Q1 & " ,Contabilizacion = " & vFldDao(Rs("Contabilizacion"))
                Q1 = Q1 & " ,TasaTef = " & vFldDao(Rs("TasaTef"))
                Q1 = Q1 & " ,TasaTex = " & vFldDao(Rs("TasaTex"))
                Q1 = Q1 & " ,Percepciones = " & vFldDao(Rs("Percepciones"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IDPerc"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE P"
    Q1 = Q1 & " SET P.IdComp = ISNULL(C.IdComp,P.IdComp),"
    Q1 = Q1 & "     p.IdCuenta = ISNULL(CU.IdCuenta, p.IdCuenta)"
    Q1 = Q1 & " FROM Percepciones P"
    Q1 = Q1 & " LEFT JOIN Comprobante C ON C.IdTras = P.IdComp AND C.IdEmpresa = P.IdEmpresa AND C.Ano = P.Ano"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = P.IdCuenta AND CU.IdEmpresa = P.IdEmpresa AND CU.Ano = P.Ano"
    Q1 = Q1 & " WHERE p.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   P.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasParamRazon(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = "DELETE FROM ParamRazon WHERE IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,IdRazon"
   Q1 = Q1 & " ,CantDias"
   Q1 = Q1 & " From ParamRazon"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)
   
   Do While Rs.EOF = False
           
        'Q1 = " SET IDENTITY_INSERT ParamRazon ON "
        Q1 = Q1 & " INSERT INTO ParamRazon"
        Q1 = Q1 & " (IdEmpresa"
        Q1 = Q1 & " ,IdRazon"
        Q1 = Q1 & " ,CantDias)"
        Q1 = Q1 & " Values"
        Q1 = Q1 & " (" & IdEmpresa
        Q1 = Q1 & " ," & vFldDao(Rs("IdRazon"))
        Q1 = Q1 & " ," & vFldDao(Rs("CantDias")) & ")"
        'Q1 = Q1 & " SET IDENTITY_INSERT ParamRazon OFF  "
        Call ExecSQL(DBSql, Q1)

      Rs.MoveNext
   Loop
   Call CloseRs(Rs)


End Sub

Public Sub TrasParamEmpresa(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " ,ValorOld"
   Q1 = Q1 & " FROM ParamEmpresa"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Q1 = Q1 & " AND Tipo NOT IN ('INITAÑO') "
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From ParamEmpresa"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND Tipo = '" & vFldDao(Rs("Tipo")) & "'"
            Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
           
            If Rs1.EOF = True Then
            
            
                'Q1 = " SET IDENTITY_INSERT ParamEmpresa ON "
                Q1 = Q1 & " INSERT INTO ParamEmpresa"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,ValorOld)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Tipo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Codigo"))
                'Q1 = Q1 & " ,'" & str(vFmt(vFldDao(Rs("Valor")))) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Valor")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("ValorOld")) & "')"
                'Q1 = Q1 & " SET IDENTITY_INSERT ParamEmpresa OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE ParamEmpresa"
                Q1 = Q1 & " SET Valor = '" & str(vFmt(vFldDao(Rs("Valor")))) & "'"
                Q1 = Q1 & " ,ValorOld = '" & vFldDao(Rs("ValorOld")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND Tipo = '" & vFldDao(Rs("Tipo")) & "'"
                Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasParam(DBSql As ADODB.Connection, DbAccess As Database)

   Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM Param"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Param"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT Tipo"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " ,Diminutivo"
   Q1 = Q1 & " ,Atributo"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Param) as Cant"
   Q1 = Q1 & " From Param"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Param"
            Q1 = Q1 & " WHERE Tipo = '" & vFldDao(Rs("Tipo")) & "'"
            Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Param ON "
                Q1 = " INSERT INTO Param"
                Q1 = Q1 & " (Tipo"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,Diminutivo"
                Q1 = Q1 & " ,Atributo)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("Tipo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Codigo"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Valor")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Atributo")) & "')"
                'Q1 = Q1 & " SET IDENTITY_INSERT Param OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Param"
                Q1 = Q1 & " SET Valor = '" & vFldDao(Rs("Valor")) & "'"
                Q1 = Q1 & " ,Diminutivo = '" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,Atributo = '" & vFldDao(Rs("Atributo")) & "'"
                Q1 = Q1 & " WHERE Tipo = '" & vFldDao(Rs("Tipo")) & "'"
                Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasNotas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = Q1 & " DELETE FROM Notas WHERE IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT Tipo"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Nota"
   Q1 = Q1 & " ,Incluir"
   Q1 = Q1 & " ,IncluirInfo"
   Q1 = Q1 & " From Notas"
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
      'Q1 = " SET IDENTITY_INSERT Notas ON "
      Q1 = " INSERT INTO Notas"
      Q1 = Q1 & " (Tipo"
      Q1 = Q1 & " ,IdEmpresa"
      Q1 = Q1 & " ,Nota"
      Q1 = Q1 & " ,Incluir"
      Q1 = Q1 & " ,IncluirInfo)"
      Q1 = Q1 & " Values"
      Q1 = Q1 & " ('" & vFldDao(Rs("Tipo")) & "'"
      Q1 = Q1 & " ," & IdEmpresa
      Q1 = Q1 & " ,'" & vFldDao(Rs("Nota")) & "'"
      Q1 = Q1 & " ," & vFldDao(Rs("Incluir"))
      Q1 = Q1 & " ," & vFldDao(Rs("IncluirInfo")) & ")"
      'Q1 = Q1 & " SET IDENTITY_INSERT Notas OFF  "
      Call ExecSQL(DBSql, Q1)

      Rs.MoveNext
   Loop
   Call CloseRs(Rs)


End Sub

Public Sub TrasMovComprobante(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'MovComprobante' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE MovComprobante ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdMov"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,IdDoc"
   Q1 = Q1 & " ,Orden"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,idCCosto"
   Q1 = Q1 & " ,idAreaNeg"
   Q1 = Q1 & " ,IdCartola"
   Q1 = Q1 & " ,DeCentraliz"
   Q1 = Q1 & " ,DePago"
   Q1 = Q1 & " ,DeRemu"
   Q1 = Q1 & " ,Nota"
   Q1 = Q1 & " ,IdDocCuota"
   Q1 = Q1 & " From MovComprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From MovComprobante"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMov"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT MovComprobante ON "
                Q1 = " INSERT INTO MovComprobante"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,IdDoc"
                Q1 = Q1 & " ,Orden"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,idCCosto"
                Q1 = Q1 & " ,idAreaNeg"
                Q1 = Q1 & " ,IdCartola"
                Q1 = Q1 & " ,DeCentraliz"
                Q1 = Q1 & " ,DePago"
                Q1 = Q1 & " ,DeRemu"
                Q1 = Q1 & " ,Nota"
                Q1 = Q1 & " ,IdDocCuota"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("idCCosto"))
                Q1 = Q1 & " ," & vFldDao(Rs("idAreaNeg"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCartola"))
                Q1 = Q1 & " ," & vFldDao(Rs("DeCentraliz"))
                Q1 = Q1 & " ," & vFldDao(Rs("DePago"))
                Q1 = Q1 & " ," & vFldDao(Rs("DeRemu"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nota")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdDocCuota"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdMov")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT MovComprobante OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE MovComprobante"
                Q1 = Q1 & " SET IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,IdDoc = " & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ,Orden = " & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,Glosa = '" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
                Q1 = Q1 & " ,idCCosto = " & vFldDao(Rs("idCCosto"))
                Q1 = Q1 & " ,idAreaNeg = " & vFldDao(Rs("idAreaNeg"))
                Q1 = Q1 & " ,IdCartola = " & vFldDao(Rs("IdCartola"))
                Q1 = Q1 & " ,DeCentraliz = " & vFldDao(Rs("DeCentraliz"))
                Q1 = Q1 & " ,DePago = " & vFldDao(Rs("DePago"))
                Q1 = Q1 & " ,DeRemu = " & vFldDao(Rs("DeRemu"))
                Q1 = Q1 & " ,Nota = '" & vFldDao(Rs("Nota")) & "'"
                Q1 = Q1 & " ,IdDocCuota = " & vFldDao(Rs("IdDocCuota"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMov"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE MV"
    Q1 = Q1 & " SET    MV.IdComp = ISNULL(COM.IdComp,0),"
    Q1 = Q1 & "        MV.IdDoc = ISNULL(DOC.IdDoc,MV.IdDoc),"
    Q1 = Q1 & "        MV.IdCuenta = ISNULL(CU.IdCuenta,MV.IdCuenta),"
    Q1 = Q1 & "        MV.idCCosto = ISNULL(CC.IdCCosto,MV.idCCosto),"
    Q1 = Q1 & "        MV.idAreaNeg = ISNULL(AN.IdAreaNegocio,MV.idAreaNeg),"
    Q1 = Q1 & "        MV.IdCartola = ISNULL(CA.IdCartola,MV.IdCartola),"
    Q1 = Q1 & "        MV.IdDocCuota = IsNull(DOCC.IdDoc, MV.IdDocCuota)"
    Q1 = Q1 & " FROM (((((((MovComprobante MV"
    Q1 = Q1 & " INNER JOIN Comprobante COM ON COM.IdTras = MV.IdComp AND COM.IdEmpresa = MV.IdEmpresa AND COM.Ano = MV.Ano)"
    Q1 = Q1 & " LEFT JOIN Documento DOC ON DOC.IdTras = MV.IdDoc AND DOC.IdEmpresa = MV.IdEmpresa AND DOC.Ano = MV.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = MV.IdCuenta AND CU.IdEmpresa = MV.IdEmpresa AND CU.Ano = MV.Ano)"
    Q1 = Q1 & " LEFT JOIN CentroCosto CC ON CC.IdTras = MV.idCCosto AND CC.IdEmpresa = MV.IdEmpresa)"
    Q1 = Q1 & " LEFT JOIN AreaNegocio AN ON AN.IdTras = MV.idAreaNeg AND AN.IdEmpresa = MV.IdEmpresa)"
    Q1 = Q1 & " LEFT JOIN Cartola CA ON CA.IdTras = MV.IdCartola AND CA.IdEmpresa = MV.IdEmpresa AND CA.Ano = MV.Ano)"
    Q1 = Q1 & " LEFT JOIN Documento DOCC ON DOCC.IdTras = MV.IdDocCuota AND DOCC.IdEmpresa = MV.IdEmpresa AND DOCC.Ano = MV.Ano)"
    Q1 = Q1 & " WHERE MV.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND MV.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasMovActivoFijo(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'MovActivoFijo' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE MovActivoFijo ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdActFijo"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdDoc"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,IdMovComp"
   Q1 = Q1 & " ,TipoMovAF"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,Cantidad"
   Q1 = Q1 & " ,Descrip"
   Q1 = Q1 & " ,Neto"
   Q1 = Q1 & " ,IVA"
   Q1 = Q1 & " ,Cred4Porc"
   Q1 = Q1 & " ,DepNormal"
   Q1 = Q1 & " ,DepAcelerada"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,DepNormalHist"
   Q1 = Q1 & " ,DepAceleradaHist"
   Q1 = Q1 & " ,NetoVenta"
   Q1 = Q1 & " ,IVAVenta"
   Q1 = Q1 & " ,FechaVentaBaja"
   Q1 = Q1 & " ,TipoDep"
   Q1 = Q1 & " ,TipoDepHist"
   Q1 = Q1 & " ,DepAcumHist"
   Q1 = Q1 & " ,VidaUtil"
   Q1 = Q1 & " ,DepAcumFinal"
   Q1 = Q1 & " ,VidaUtilResidual"
   Q1 = Q1 & " ,FExported"
   Q1 = Q1 & " ,FechaUtilizacion"
   Q1 = Q1 & " ,NoDepreciable"
   Q1 = Q1 & " ,ValCred33"
   Q1 = Q1 & " ,ValReajustadoNeto"
   Q1 = Q1 & " ,IdActFijoOld"
   Q1 = Q1 & " ,IdActFijoOldTmp"
   Q1 = Q1 & " ,TotalmenteDepreciado"
   Q1 = Q1 & " ,ValorLibro"
   Q1 = Q1 & " ,FImported"
   Q1 = Q1 & " ,ValReajustadoNetoAnt"
   Q1 = Q1 & " ,Cred4PorcAnoInit"
   Q1 = Q1 & " ,FechaImportFile"
   Q1 = Q1 & " ,DepInstant"
   Q1 = Q1 & " ,DepDecimaParte"
   Q1 = Q1 & " ,DepInstantHist"
   Q1 = Q1 & " ,DepDecimaParteHist"
   Q1 = Q1 & " ,VidaUtilAnos"
   Q1 = Q1 & " ,TipoDepLey21210"
   Q1 = Q1 & " ,DepDecimaParte2"
   Q1 = Q1 & " ,DepDecimaParte2Hist"
   Q1 = Q1 & " ,PatenteRol"
   Q1 = Q1 & " ,NombreProy"
   Q1 = Q1 & " ,FechaProy"
   Q1 = Q1 & " ,TipoDepLey21210Hist"
   Q1 = Q1 & " ,DepLey21256"
   Q1 = Q1 & " ,DepLey21256Hist"
   If VerDBAccess > 738 Then
    Q1 = Q1 & " ,idCCosto"
    Q1 = Q1 & " ,IdAreaNeg"
   End If
   
   Q1 = Q1 & " From MovActivoFijo"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From MovActivoFijo"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdActFijo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT MovActivoFijo ON "
                Q1 = " INSERT INTO MovActivoFijo"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdDoc"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,IdMovComp"
                Q1 = Q1 & " ,TipoMovAF"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,Cantidad"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Neto"
                Q1 = Q1 & " ,IVA"
                Q1 = Q1 & " ,Cred4Porc"
                Q1 = Q1 & " ,DepNormal"
                Q1 = Q1 & " ,DepAcelerada"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,DepNormalHist"
                Q1 = Q1 & " ,DepAceleradaHist"
                Q1 = Q1 & " ,NetoVenta"
                Q1 = Q1 & " ,IVAVenta"
                Q1 = Q1 & " ,FechaVentaBaja"
                Q1 = Q1 & " ,TipoDep"
                Q1 = Q1 & " ,TipoDepHist"
                Q1 = Q1 & " ,DepAcumHist"
                Q1 = Q1 & " ,VidaUtil"
                Q1 = Q1 & " ,DepAcumFinal"
                Q1 = Q1 & " ,VidaUtilResidual"
                Q1 = Q1 & " ,FExported"
                Q1 = Q1 & " ,FechaUtilizacion"
                Q1 = Q1 & " ,NoDepreciable"
                Q1 = Q1 & " ,ValCred33"
                Q1 = Q1 & " ,ValReajustadoNeto"
                Q1 = Q1 & " ,IdActFijoOld"
                Q1 = Q1 & " ,IdActFijoOldTmp"
                Q1 = Q1 & " ,TotalmenteDepreciado"
                Q1 = Q1 & " ,ValorLibro"
                Q1 = Q1 & " ,FImported"
                Q1 = Q1 & " ,ValReajustadoNetoAnt"
                Q1 = Q1 & " ,Cred4PorcAnoInit"
                Q1 = Q1 & " ,FechaImportFile"
                Q1 = Q1 & " ,DepInstant"
                Q1 = Q1 & " ,DepDecimaParte"
                Q1 = Q1 & " ,DepInstantHist"
                Q1 = Q1 & " ,DepDecimaParteHist"
                Q1 = Q1 & " ,VidaUtilAnos"
                Q1 = Q1 & " ,TipoDepLey21210"
                Q1 = Q1 & " ,DepDecimaParte2"
                Q1 = Q1 & " ,DepDecimaParte2Hist"
                Q1 = Q1 & " ,PatenteRol"
                Q1 = Q1 & " ,NombreProy"
                Q1 = Q1 & " ,FechaProy"
                Q1 = Q1 & " ,TipoDepLey21210Hist"
                Q1 = Q1 & " ,DepLey21256"
                Q1 = Q1 & " ,DepLey21256Hist"
                If VerDBAccess > 738 Then
                    Q1 = Q1 & " ,idCCosto"
                    Q1 = Q1 & " ,IdAreaNeg"
                End If
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdMovComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoMovAF"))
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & vFldDao(Rs("Cantidad"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Neto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("Cred4Porc"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepNormal"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAcelerada"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepNormalHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAceleradaHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("NetoVenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVAVenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaVentaBaja"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDep"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDepHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAcumHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtil"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAcumFinal"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtilResidual"))
                Q1 = Q1 & " ," & vFldDao(Rs("FExported"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaUtilizacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("NoDepreciable"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValCred33"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValReajustadoNeto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdActFijoOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdActFijoOldTmp"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotalmenteDepreciado"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValorLibro"))
                Q1 = Q1 & " ," & vFldDao(Rs("FImported"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValReajustadoNetoAnt"))
                Q1 = Q1 & " ," & vFldDao(Rs("Cred4PorcAnoInit"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaImportFile"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepInstant"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepDecimaParte"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepInstantHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepDecimaParteHist"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtilAnos"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDepLey21210"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepDecimaParte2"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepDecimaParte2Hist"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("PatenteRol")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombreProy")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("FechaProy"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDepLey21210Hist"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepLey21256"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepLey21256Hist"))
                If VerDBAccess > 738 Then
                    Q1 = Q1 & " ," & vFldDao(Rs("idCCosto"))
                    Q1 = Q1 & " ," & vFldDao(Rs("IdAreaNeg"))
                End If
                Q1 = Q1 & " ," & vFldDao(Rs("IdActFijo")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT MovActivoFijo OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE MovActivoFijo"
                Q1 = Q1 & " SET IdDoc = " & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ,IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,IdMovComp = " & vFldDao(Rs("IdMovComp"))
                Q1 = Q1 & " ,TipoMovAF = " & vFldDao(Rs("TipoMovAF"))
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Cantidad = " & vFldDao(Rs("Cantidad"))
                Q1 = Q1 & " ,Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,Neto = " & vFldDao(Rs("Neto"))
                Q1 = Q1 & " ,IVA = " & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ,Cred4Porc = " & vFldDao(Rs("Cred4Porc"))
                Q1 = Q1 & " ,DepNormal = " & vFldDao(Rs("DepNormal"))
                Q1 = Q1 & " ,DepAcelerada = " & vFldDao(Rs("DepAcelerada"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,DepNormalHist = " & vFldDao(Rs("DepNormalHist"))
                Q1 = Q1 & " ,DepAceleradaHist = " & vFldDao(Rs("DepAceleradaHist"))
                Q1 = Q1 & " ,NetoVenta = " & vFldDao(Rs("NetoVenta"))
                Q1 = Q1 & " ,IVAVenta = " & vFldDao(Rs("IVAVenta"))
                Q1 = Q1 & " ,FechaVentaBaja = " & vFldDao(Rs("FechaVentaBaja"))
                Q1 = Q1 & " ,TipoDep = " & vFldDao(Rs("TipoDep"))
                Q1 = Q1 & " ,TipoDepHist = " & vFldDao(Rs("TipoDepHist"))
                Q1 = Q1 & " ,DepAcumHist = " & vFldDao(Rs("DepAcumHist"))
                Q1 = Q1 & " ,VidaUtil = " & vFldDao(Rs("VidaUtil"))
                Q1 = Q1 & " ,DepAcumFinal = " & vFldDao(Rs("DepAcumFinal"))
                Q1 = Q1 & " ,VidaUtilResidual = " & vFldDao(Rs("VidaUtilResidual"))
                Q1 = Q1 & " ,FExported = " & vFldDao(Rs("FExported"))
                Q1 = Q1 & " ,FechaUtilizacion = " & vFldDao(Rs("FechaUtilizacion"))
                Q1 = Q1 & " ,NoDepreciable = " & vFldDao(Rs("NoDepreciable"))
                Q1 = Q1 & " ,ValCred33 = " & vFldDao(Rs("ValCred33"))
                Q1 = Q1 & " ,ValReajustadoNeto = " & vFldDao(Rs("ValReajustadoNeto"))
                Q1 = Q1 & " ,IdActFijoOld = " & vFldDao(Rs("IdActFijoOld"))
                Q1 = Q1 & " ,IdActFijoOldTmp = " & vFldDao(Rs("IdActFijoOldTmp"))
                Q1 = Q1 & " ,TotalmenteDepreciado = " & vFldDao(Rs("TotalmenteDepreciado"))
                Q1 = Q1 & " ,ValorLibro = " & vFldDao(Rs("ValorLibro"))
                Q1 = Q1 & " ,FImported = " & vFldDao(Rs("FImported"))
                Q1 = Q1 & " ,ValReajustadoNetoAnt = " & vFldDao(Rs("ValReajustadoNetoAnt"))
                Q1 = Q1 & " ,Cred4PorcAnoInit = " & vFldDao(Rs("Cred4PorcAnoInit"))
                Q1 = Q1 & " ,FechaImportFile = " & vFldDao(Rs("FechaImportFile"))
                Q1 = Q1 & " ,DepInstant = " & vFldDao(Rs("DepInstant"))
                Q1 = Q1 & " ,DepDecimaParte = " & vFldDao(Rs("DepDecimaParte"))
                Q1 = Q1 & " ,DepInstantHist = " & vFldDao(Rs("DepInstantHist"))
                Q1 = Q1 & " ,DepDecimaParteHist = " & vFldDao(Rs("DepDecimaParteHist"))
                Q1 = Q1 & " ,VidaUtilAnos = " & vFldDao(Rs("VidaUtilAnos"))
                Q1 = Q1 & " ,TipoDepLey21210 = " & vFldDao(Rs("TipoDepLey21210"))
                Q1 = Q1 & " ,DepDecimaParte2 = " & vFldDao(Rs("DepDecimaParte2"))
                Q1 = Q1 & " ,DepDecimaParte2Hist = " & vFldDao(Rs("DepDecimaParte2Hist"))
                Q1 = Q1 & " ,PatenteRol = '" & vFldDao(Rs("PatenteRol")) & "'"
                Q1 = Q1 & " ,NombreProy = '" & vFldDao(Rs("NombreProy")) & "'"
                Q1 = Q1 & " ,FechaProy = " & vFldDao(Rs("FechaProy"))
                Q1 = Q1 & " ,TipoDepLey21210Hist = " & vFldDao(Rs("TipoDepLey21210Hist"))
                Q1 = Q1 & " ,DepLey21256 = " & vFldDao(Rs("DepLey21256"))
                Q1 = Q1 & " ,DepLey21256Hist = " & vFldDao(Rs("DepLey21256Hist"))
                If VerDBAccess > 738 Then
                    Q1 = Q1 & " ,idCCosto = " & vFldDao(Rs("idCCosto"))
                    Q1 = Q1 & " ,IdAreaNeg = " & vFldDao(Rs("IdAreaNeg"))
                End If
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdActFijo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE MAC"
    Q1 = Q1 & " SET    MAC.IdDoc = ISNULL(D.IdDoc,MAC.IdDoc),"
    Q1 = Q1 & "        MAC.IdComp = ISNULL(C.IdComp,MAC.IdComp),"
    Q1 = Q1 & "        MAC.IdMovComp = ISNULL(MC.IdMov,MAC.IdMovComp),"
    Q1 = Q1 & "        MAC.IdCuenta = ISNULL(CU.idCuenta,MAC.IdCuenta),"
    Q1 = Q1 & "        MAC.IdActFijoOld = ISNULL(MACO.IdDoc,MAC.IdActFijoOld),"
    Q1 = Q1 & "        MAC.idCCosto = ISNULL(CC.IdCCosto,MAC.idCCosto),"
    Q1 = Q1 & "        Mac.IdAreaNeg = IsNull(AN.IdAreaNegocio, Mac.IdAreaNeg)"
    Q1 = Q1 & " FROM (((((((MovActivoFijo MAC"
    Q1 = Q1 & " LEFT JOIN Documento D ON D.IdTras = MAC.IdDoc AND D.IdEmpresa = MAC.IdEmpresa AND D.Ano = MAC.Ano)"
    Q1 = Q1 & " LEFT JOIN Comprobante C ON C.IdTras = MAC.IdComp AND C.IdEmpresa = MAC.IdEmpresa AND C.Ano = MAC.Ano)"
    Q1 = Q1 & " LEFT JOIN MovComprobante MC ON MC.IdTras = MAC.IdMovComp AND MC.IdEmpresa = MAC.IdEmpresa AND MC.Ano = MAC.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = MAC.IdCuenta AND CU.IdEmpresa = MAC.IdEmpresa AND CU.Ano = MAC.Ano)"
    Q1 = Q1 & " LEFT JOIN MovActivoFijo MACO ON MACO.IdTras = MAC.IdActFijoOld AND MACO.IdEmpresa = MAC.IdEmpresa AND MACO.Ano = MAC.Ano - 1)"
    Q1 = Q1 & " LEFT JOIN CentroCosto CC ON CC.IdTras = MAC.idCCosto  AND CC.IdEmpresa = MAC.IdEmpresa)"
    Q1 = Q1 & " LEFT JOIN AreaNegocio AN ON AN.IdTras = MAC.IdAreaNeg AND AN.IdEmpresa = MAC.IdEmpresa)"
    Q1 = Q1 & " WHERE Mac.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND MAC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)
    
    Q1 = " SELECT MA.IdActFijoOld, ISNULL(MOA.IdActFijo,MA.IdActFijoOld),"
    Q1 = Q1 & " MA.IdActFijoOldTmp , ISNULL(MOAC.IdActFijo, MA.IdActFijoOldTmp)"
    Q1 = Q1 & " FROM MovActivoFijo MA"
    Q1 = Q1 & " LEFT JOIN MovActivoFijo MOA ON MOA.IdTras = MA.IdActFijoOld AND MOA.IdEmpresa = MA.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN MovActivoFijo MOAC ON MOAC.IdTras = MA.IdActFijoOldTmp AND MOAC.IdEmpresa = MA.IdEmpresa"
    Q1 = Q1 & " WHERE MA.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   MA.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasMonedas(DBSql As ADODB.Connection, DbAccess As Database)
   
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Monedas"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)
   
   Q1 = "SELECT idMoneda"
   Q1 = Q1 & " ,Descrip"
   Q1 = Q1 & " ,Simbolo"
   Q1 = Q1 & " ,DecInf"
   Q1 = Q1 & " ,DecVenta"
   Q1 = Q1 & " ,Caracteristica"
   Q1 = Q1 & " ,CodAduana"
   Q1 = Q1 & " ,EsFijo"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Monedas) as Cant"
   Q1 = Q1 & " From Monedas"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Monedas"
            Q1 = Q1 & " WHERE idMoneda = " & vFldDao(Rs("idMoneda"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                'Q1 = " SET IDENTITY_INSERT Monedas ON "
                Q1 = " INSERT INTO Monedas"
                Q1 = Q1 & " (idMoneda"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Simbolo"
                Q1 = Q1 & " ,DecInf"
                Q1 = Q1 & " ,DecVenta"
                Q1 = Q1 & " ,Caracteristica"
                Q1 = Q1 & " ,CodAduana"
                Q1 = Q1 & " ,EsFijo)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idMoneda"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Simbolo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("DecInf"))
                Q1 = Q1 & " ," & vFldDao(Rs("DecVenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("Caracteristica"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodAduana")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("EsFijo")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Monedas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Monedas"
                Q1 = Q1 & " SET Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,Simbolo = '" & vFldDao(Rs("Simbolo")) & "'"
                Q1 = Q1 & " ,DecInf = " & vFldDao(Rs("DecInf"))
                Q1 = Q1 & " ,DecVenta = " & vFldDao(Rs("DecVenta"))
                Q1 = Q1 & " ,Caracteristica = " & vFldDao(Rs("Caracteristica"))
                Q1 = Q1 & " ,CodAduana = '" & vFldDao(Rs("CodAduana")) & "'"
                Q1 = Q1 & " ,EsFijo = " & vFldDao(Rs("EsFijo"))
                Q1 = Q1 & " WHERE idMoneda = " & vFldDao(Rs("idMoneda"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasMembrete(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = Q1 & " DELETE FROM Membrete WHERE IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT TituloMembrete1"
   Q1 = Q1 & " ,TituloMembrete2"
   Q1 = Q1 & " ,Texto1"
   Q1 = Q1 & " ,Texto2"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " From Membrete"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
           
                'Q1 = " SET IDENTITY_INSERT Membrete ON "
                Q1 = " INSERT INTO Membrete"
                Q1 = Q1 & " (TituloMembrete1"
                Q1 = Q1 & " ,TituloMembrete2"
                Q1 = Q1 & " ,Texto1"
                Q1 = Q1 & " ,Texto2"
                Q1 = Q1 & " ,IdEmpresa)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("TituloMembrete1")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("TituloMembrete2")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Texto1")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Texto2")) & "'"
                Q1 = Q1 & " ," & IdEmpresa & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Membrete OFF  "
                Call ExecSQL(DBSql, Q1)

      Rs.MoveNext
   Loop
   Call CloseRs(Rs)


End Sub

Public Sub TrasLParam()

'   Q1 = "SELECT Codigo"
'   Q1 = Q1 & " ,Valor"
'   Q1 = Q1 & " From lParam"
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT LParam ON "
'      Q1 = Q1 & " INSERT INTO LParam"
'      Q1 = Q1 & " (Codigo"
'      Q1 = Q1 & " ,Valor)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("Codigo"))
'      Q1 = Q1 & " ,'" & vFld(Rs("Valor")) & "')"
'      Q1 = Q1 & " SET IDENTITY_INSERT LParam OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)


End Sub

Public Sub TrasLogImpreso()

'   Q1 = "SELECT IdLog"
'   Q1 = Q1 & " ,IdEmpresa"
'   Q1 = Q1 & " ,Ano"
'   Q1 = Q1 & " ,Fecha"
'   Q1 = Q1 & " ,CorrInicio"
'   Q1 = Q1 & " ,CorrFin"
'   Q1 = Q1 & " ,idInforme"
'   Q1 = Q1 & " ,Estado"
'   Q1 = Q1 & " ,Comentario"
'   Q1 = Q1 & " ,IdUsuario"
'   Q1 = Q1 & " ,Mes"
'   Q1 = Q1 & " ,FDesde"
'   Q1 = Q1 & " ,FHasta"
'   Q1 = Q1 & " From LogImpreso"
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT LogImpreso ON "
'      Q1 = Q1 & " INSERT INTO LogImpreso"
'      Q1 = Q1 & " (IdEmpresa"
'      Q1 = Q1 & " ,Ano"
'      Q1 = Q1 & " ,Fecha"
'      Q1 = Q1 & " ,CorrInicio"
'      Q1 = Q1 & " ,CorrFin"
'      Q1 = Q1 & " ,idInforme"
'      Q1 = Q1 & " ,Estado"
'      Q1 = Q1 & " ,Comentario"
'      Q1 = Q1 & " ,IdUsuario"
'      Q1 = Q1 & " ,Mes"
'      Q1 = Q1 & " ,FDesde"
'      Q1 = Q1 & " ,FHasta)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("IdEmpresa"))
'      Q1 = Q1 & " ," & vFld(Rs("Ano"))
'      Q1 = Q1 & " ," & vFld(Rs("Fecha"))
'      Q1 = Q1 & " ," & vFld(Rs("CorrInicio"))
'      Q1 = Q1 & " ," & vFld(Rs("CorrFin"))
'      Q1 = Q1 & " ," & vFld(Rs("idInforme"))
'      Q1 = Q1 & " ," & vFld(Rs("Estado"))
'      Q1 = Q1 & " ,'" & vFld(Rs("Comentario")) & "'"
'      Q1 = Q1 & " ," & vFld(Rs("IdUsuario"))
'      Q1 = Q1 & " ," & vFld(Rs("Mes"))
'      Q1 = Q1 & " ," & vFld(Rs("FDesde"))
'      Q1 = Q1 & " ," & vFld(Rs("FHasta")) & ")"
'      Q1 = Q1 & " SET IDENTITY_INSERT LogImpreso OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)


End Sub

Public Sub TrasLogComprobantes(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'LogComprobantes' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE LogComprobantes ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdLog"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,IdUsuario"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,IdOper"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,CorrComp"
   Q1 = Q1 & " ,FechaComp"
   Q1 = Q1 & " ,TipoComp"
   Q1 = Q1 & " ,EstadoComp"
   Q1 = Q1 & " ,TipoAjusteComp"
   Q1 = Q1 & " From LogComprobantes"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From LogComprobantes"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdLog"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT LogComprobantes ON "
                Q1 = " INSERT INTO LogComprobantes"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,IdUsuario"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,IdOper"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,CorrComp"
                Q1 = Q1 & " ,FechaComp"
                Q1 = Q1 & " ,TipoComp"
                Q1 = Q1 & " ,EstadoComp"
                Q1 = Q1 & " ,TipoAjusteComp"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Fecha"))))
                Q1 = Q1 & " ," & vFldDao(Rs("IdOper"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("CorrComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("EstadoComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoAjusteComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdLog")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT LogComprobantes OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE LogComprobantes"
                Q1 = Q1 & " SET IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,Fecha = " & str(vFmt(vFldDao(Rs("Fecha"))))
                Q1 = Q1 & " ,IdOper = " & vFldDao(Rs("IdOper"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,CorrComp = " & vFldDao(Rs("CorrComp"))
                Q1 = Q1 & " ,FechaComp = " & vFldDao(Rs("FechaComp"))
                Q1 = Q1 & " ,TipoComp = " & vFldDao(Rs("TipoComp"))
                Q1 = Q1 & " ,EstadoComp = " & vFldDao(Rs("EstadoComp"))
                Q1 = Q1 & " ,TipoAjusteComp = " & vFldDao(Rs("TipoAjusteComp"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdLog"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE LC"
    Q1 = Q1 & " SET LC.idcomp = ISNULL(C.idcomp, LC.idcomp)"
    Q1 = Q1 & " FROM LogComprobantes LC"
    Q1 = Q1 & " LEFT JOIN Comprobante C ON C.IdTras = LC.IdComp AND C.IdEmpresa = LC.IdEmpresa AND C.Ano = LC.Ano"
    Q1 = Q1 & " WHERE LC.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND LC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasLockAction(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'LockAction' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE LockAction ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT Fecha"
   Q1 = Q1 & " ,idLock"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,PcName"
   Q1 = Q1 & " ,hInstance"
   Q1 = Q1 & " ,idAction"
   Q1 = Q1 & " ,idItem"
   Q1 = Q1 & " From LockAction"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From LockAction"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idLock"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT LockAction ON "
                Q1 = " INSERT INTO LockAction"
                Q1 = Q1 & " (Fecha"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,PcName"
                Q1 = Q1 & " ,hInstance"
                Q1 = Q1 & " ,idAction"
                Q1 = Q1 & " ,idItem"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("Fecha")) & "'"
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("PcName")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("hInstance"))
                Q1 = Q1 & " ," & vFldDao(Rs("idAction"))
                Q1 = Q1 & " ," & vFldDao(Rs("idItem"))
                Q1 = Q1 & " ," & vFldDao(Rs("idLock")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT LockAction OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE LockAction"
                Q1 = Q1 & " SET Fecha = '" & vFldDao(Rs("Fecha")) & "'"
                Q1 = Q1 & " ,PcName = '" & vFldDao(Rs("PcName")) & "'"
                Q1 = Q1 & " ,hInstance = " & vFldDao(Rs("hInstance"))
                Q1 = Q1 & " ,idAction = " & vFldDao(Rs("idAction"))
                Q1 = Q1 & " ,idItem = " & vFldDao(Rs("idItem"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idLock"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasLibroCaja(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'LibroCaja' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE LibroCaja ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdLibroCaja"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdDoc"
   Q1 = Q1 & " ,TipoOper"
   Q1 = Q1 & " ,TipoDoc"
   Q1 = Q1 & " ,TipoLib"
   Q1 = Q1 & " ,NumDoc"
   Q1 = Q1 & " ,NumDocHasta"
   Q1 = Q1 & " ,DTE"
   Q1 = Q1 & " ,IdEntidad"
   Q1 = Q1 & " ,RutEntidad"
   Q1 = Q1 & " ,NombreEntidad"
   Q1 = Q1 & " ,FechaOperacion"
   Q1 = Q1 & " ,Afecto"
   Q1 = Q1 & " ,IVA"
   Q1 = Q1 & " ,Exento"
   Q1 = Q1 & " ,OtroImp"
   Q1 = Q1 & " ,Total"
   Q1 = Q1 & " ,Pagado"
   Q1 = Q1 & " ,Descrip"
   Q1 = Q1 & " ,ConEntRel"
   Q1 = Q1 & " ,OperDevengada"
   Q1 = Q1 & " ,PagoAPlazo"
   Q1 = Q1 & " ,FechaExigPago"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,IdUsuario"
   Q1 = Q1 & " ,FechaCreacion"
   Q1 = Q1 & " ,IVAIrrec"
   Q1 = Q1 & " ,FechaIngresoLibro"
   Q1 = Q1 & " ,IdEntReal"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,Ingreso"
   Q1 = Q1 & " ,Egreso"
   Q1 = Q1 & " ,MontoAfectaBaseImp"
   Q1 = Q1 & " From LibroCaja"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From LibroCaja"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdLibroCaja"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT LibroCaja ON "
                Q1 = " INSERT INTO LibroCaja"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdDoc"
                Q1 = Q1 & " ,TipoOper"
                Q1 = Q1 & " ,TipoDoc"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,NumDoc"
                Q1 = Q1 & " ,NumDocHasta"
                Q1 = Q1 & " ,DTE"
                Q1 = Q1 & " ,IdEntidad"
                Q1 = Q1 & " ,RutEntidad"
                Q1 = Q1 & " ,NombreEntidad"
                Q1 = Q1 & " ,FechaOperacion"
                Q1 = Q1 & " ,Afecto"
                Q1 = Q1 & " ,IVA"
                Q1 = Q1 & " ,Exento"
                Q1 = Q1 & " ,OtroImp"
                Q1 = Q1 & " ,Total"
                Q1 = Q1 & " ,Pagado"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,ConEntRel"
                Q1 = Q1 & " ,OperDevengada"
                Q1 = Q1 & " ,PagoAPlazo"
                Q1 = Q1 & " ,FechaExigPago"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,IdUsuario"
                Q1 = Q1 & " ,FechaCreacion"
                Q1 = Q1 & " ,IVAIrrec"
                Q1 = Q1 & " ,FechaIngresoLibro"
                Q1 = Q1 & " ,IdEntReal"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,Ingreso"
                Q1 = Q1 & " ,Egreso"
                Q1 = Q1 & " ,MontoAfectaBaseImp"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoOper"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDoc")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDocHasta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("DTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("RutEntidad")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombreEntidad")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("FechaOperacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("Afecto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("Exento"))
                Q1 = Q1 & " ," & vFldDao(Rs("OtroImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Total"))
                Q1 = Q1 & " ," & vFldDao(Rs("Pagado"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("ConEntRel"))
                Q1 = Q1 & " ," & vFldDao(Rs("OperDevengada"))
                Q1 = Q1 & " ," & vFldDao(Rs("PagoAPlazo"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaExigPago"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaCreacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVAIrrec"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaIngresoLibro"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdEntReal"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Ingreso"))
                Q1 = Q1 & " ," & vFldDao(Rs("Egreso"))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoAfectaBaseImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdLibroCaja")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT LibroCaja OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE LibroCaja"
                Q1 = Q1 & " SET IdDoc = " & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ,TipoOper = " & vFldDao(Rs("TipoOper"))
                Q1 = Q1 & " ,TipoDoc = " & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ,TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,NumDoc = '" & vFldDao(Rs("NumDoc")) & "'"
                Q1 = Q1 & " ,NumDocHasta = '" & vFldDao(Rs("NumDocHasta")) & "'"
                Q1 = Q1 & " ,DTE = " & vFldDao(Rs("DTE"))
                Q1 = Q1 & " ,IdEntidad = " & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ,RutEntidad = '" & vFldDao(Rs("RutEntidad")) & "'"
                Q1 = Q1 & " ,NombreEntidad = '" & vFldDao(Rs("NombreEntidad")) & "'"
                Q1 = Q1 & " ,FechaOperacion = " & vFldDao(Rs("FechaOperacion"))
                Q1 = Q1 & " ,Afecto = " & vFldDao(Rs("Afecto"))
                Q1 = Q1 & " ,IVA = " & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ,Exento = " & vFldDao(Rs("Exento"))
                Q1 = Q1 & " ,OtroImp = " & vFldDao(Rs("OtroImp"))
                Q1 = Q1 & " ,Total = " & vFldDao(Rs("Total"))
                Q1 = Q1 & " ,Pagado = " & vFldDao(Rs("Pagado"))
                Q1 = Q1 & " ,Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,ConEntRel = " & vFldDao(Rs("ConEntRel"))
                Q1 = Q1 & " ,OperDevengada = " & vFldDao(Rs("OperDevengada"))
                Q1 = Q1 & " ,PagoAPlazo = " & vFldDao(Rs("PagoAPlazo"))
                Q1 = Q1 & " ,FechaExigPago = " & vFldDao(Rs("FechaExigPago"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,FechaCreacion = " & vFldDao(Rs("FechaCreacion"))
                Q1 = Q1 & " ,IVAIrrec = " & vFldDao(Rs("IVAIrrec"))
                Q1 = Q1 & " ,FechaIngresoLibro = " & vFldDao(Rs("FechaIngresoLibro"))
                Q1 = Q1 & " ,IdEntReal = " & vFldDao(Rs("IdEntReal"))
                Q1 = Q1 & " ,IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,Ingreso = " & vFldDao(Rs("Ingreso"))
                Q1 = Q1 & " ,Egreso = " & vFldDao(Rs("Egreso"))
                Q1 = Q1 & " ,MontoAfectaBaseImp = " & vFldDao(Rs("MontoAfectaBaseImp"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdLibroCaja"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE LC"
    Q1 = Q1 & "  SET    LC.IdDoc = ISNULL(D.IdDoc,LC.IdDoc),"
    Q1 = Q1 & "         LC.IdEntidad = ISNULL(E.IdEntidad,LC.IdEntidad),"
    Q1 = Q1 & "         LC.IdEntReal = ISNULL(E.IdEntidad,LC.IdEntReal),"
    Q1 = Q1 & "  LC.idcomp = ISNULL(c.idcomp, lc.idcomp)"
    Q1 = Q1 & " FROM LibroCaja LC"
    Q1 = Q1 & " LEFT JOIN Documento D ON D.IdTras = LC.IdDoc AND D.IdEmpresa = LC.IdEmpresa AND D.Ano = LC.Ano"
    Q1 = Q1 & " LEFT JOIN Entidades E ON E.IdTras = LC.IdEntidad AND E.IdEmpresa = LC.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN Entidades EN ON EN.IdTras = LC.IdEntReal AND EN.IdEmpresa = LC.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN Comprobante C ON C.IdTras = LC.IdComp AND C.IdEmpresa = LC.IdEmpresa AND C.Ano = LC.Ano"
    Q1 = Q1 & " WHERE lc.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND LC.Ano = " & Ano

    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasIPC(DBSql As ADODB.Connection, DbAccess As Database)
   
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From IPC"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT AnoMes"
   Q1 = Q1 & " ,pIPC"
   Q1 = Q1 & " ,vIPC"
   Q1 = Q1 & " ,fCM"
   Q1 = Q1 & " ,aIPC"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM IPC) as Cant"
   Q1 = Q1 & " From IPC"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From IPC"
            Q1 = Q1 & " WHERE AnoMes = " & vFldDao(Rs("AnoMes"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                'Q1 = " SET IDENTITY_INSERT IPC ON "
                Q1 = " INSERT INTO IPC"
                Q1 = Q1 & " (AnoMes"
                Q1 = Q1 & " ,pIPC"
                Q1 = Q1 & " ,vIPC"
                Q1 = Q1 & " ,fCM"
                Q1 = Q1 & " ,aIPC)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("AnoMes"))
                Q1 = Q1 & " ," & Replace(vFldDao(Rs("pIPC")), ",", ".")
                Q1 = Q1 & " ," & Replace(vFldDao(Rs("vIPC")), ",", ".")
                Q1 = Q1 & " ," & Replace(vFldDao(Rs("fCM")), ",", ".")
                Q1 = Q1 & " ," & Replace(vFldDao(Rs("aIPC")), ",", ".") & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT IPC OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE IPC"
                Q1 = Q1 & " SET pIPC = '" & Replace(vFldDao(Rs("pIPC")), ",", ".") & "'"
                Q1 = Q1 & " ,vIPC = '" & Replace(vFldDao(Rs("vIPC")), ",", ".") & "'"
                Q1 = Q1 & " ,fCM = '" & Replace(vFldDao(Rs("fCM")), ",", ".") & "'"
                Q1 = Q1 & " ,aIPC = '" & Replace(vFldDao(Rs("aIPC")), ",", ".") & "'"
                Q1 = Q1 & " WHERE AnoMes = " & vFldDao(Rs("AnoMes"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasInfoAnualDJ1847(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

   Q1 = "SELECT IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdEntSupervisora"
   Q1 = Q1 & " ,AnoAjusteIFRS"
   Q1 = Q1 & " ,FolioInicial"
   Q1 = Q1 & " ,FolioFinal"
   Q1 = Q1 & " ,IdAjustesRLI"
   Q1 = Q1 & " From InfoAnualDJ1847"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From InfoAnualDJ1847"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            'Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT InfoAnualDJ1847 ON "
                Q1 = " INSERT INTO InfoAnualDJ1847"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdEntSupervisora"
                Q1 = Q1 & " ,AnoAjusteIFRS"
                Q1 = Q1 & " ,FolioInicial"
                Q1 = Q1 & " ,FolioFinal"
                Q1 = Q1 & " ,IdAjustesRLI)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdEntSupervisora"))
                Q1 = Q1 & " ," & vFldDao(Rs("AnoAjusteIFRS"))
                Q1 = Q1 & " ," & vFldDao(Rs("FolioInicial"))
                Q1 = Q1 & " ," & vFldDao(Rs("FolioFinal"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAjustesRLI")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT InfoAnualDJ1847 OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE InfoAnualDJ1847"
                Q1 = Q1 & " SET IdEntSupervisora = " & vFldDao(Rs("IdEntSupervisora"))
                Q1 = Q1 & " ,AnoAjusteIFRS = " & vFldDao(Rs("AnoAjusteIFRS"))
                Q1 = Q1 & " ,FolioInicial = " & vFldDao(Rs("FolioInicial"))
                Q1 = Q1 & " ,FolioFinal = " & vFldDao(Rs("FolioFinal"))
                Q1 = Q1 & " ,IdAjustesRLI = " & vFldDao(Rs("IdAjustesRLI"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                'Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasImpuestos(DBSql As ADODB.Connection, DbAccess As Database)

   Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From Impuestos"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idImpuesto"
   Q1 = Q1 & " ,Impuesto"
   Q1 = Q1 & " ,Porcentaje"
   Q1 = Q1 & " ,FechaDesde"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Impuestos) as Cant"
   Q1 = Q1 & " From Impuestos"
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
   
   If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Impuestos"
            Q1 = Q1 & " WHERE idImpuesto = " & vFldDao(Rs("idImpuesto"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT Impuestos ON "
                Q1 = Q1 & " INSERT INTO Impuestos"
                Q1 = Q1 & " (idImpuesto"
                Q1 = Q1 & " ,Impuesto"
                Q1 = Q1 & " ,Porcentaje"
                Q1 = Q1 & " ,FechaDesde)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idImpuesto"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Impuesto")) & "'"
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Porcentaje"))))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaDesde")) & ")"
                Q1 = Q1 & " SET IDENTITY_INSERT Impuestos OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Impuestos"
                Q1 = Q1 & " SET Impuesto = '" & vFldDao(Rs("Impuesto")) & "'"
                Q1 = Q1 & " ,Porcentaje = " & str(vFmt(vFldDao(Rs("Porcentaje"))))
                Q1 = Q1 & " ,FechaDesde = " & vFldDao(Rs("FechaDesde"))
                Q1 = Q1 & " WHERE idImpuesto = " & vFldDao(Rs("idImpuesto"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasImpAdic(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'ImpAdic' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE ImpAdic ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdImpAdic"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,TipoLib"
   Q1 = Q1 & " ,TipoValor"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,Tasa"
   Q1 = Q1 & " ,EsRecuperable"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " From ImpAdic"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From ImpAdic"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdImpAdic"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT ImpAdic ON "
                Q1 = " INSERT INTO ImpAdic"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,TipoValor"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,Tasa"
                Q1 = Q1 & " ,EsRecuperable"
                Q1 = Q1 & " ,CodCuenta"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoValor"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tasa"))
                Q1 = Q1 & " ," & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdImpAdic")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT ImpAdic OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE ImpAdic"
                Q1 = Q1 & " SET TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,TipoValor = " & vFldDao(Rs("TipoValor"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,Tasa = " & vFldDao(Rs("Tasa"))
                Q1 = Q1 & " ,EsRecuperable = " & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdImpAdic"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE IA"
    Q1 = Q1 & " SET IA.IdCuenta = IsNull(CU.IdCuenta, IA.IdCuenta)"
    Q1 = Q1 & " FROM ImpAdic IA"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = IA.IdCuenta AND CU.IdEmpresa = IA.IdEmpresa AND CU.Ano = IA.Ano"
    Q1 = Q1 & " WHERE IA.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND IA.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasIFRS_PlanIFRS(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(idCuenta) as Cant"
    Q1 = Q1 & " From IFRS_PlanIFRS"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idCuenta"
   Q1 = Q1 & " ,idPadre"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,Nivel"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,TipoCapPropio"
   Q1 = Q1 & " ,CodF22"
   Q1 = Q1 & " ,Atrib1"
   Q1 = Q1 & " ,Atrib2"
   Q1 = Q1 & " ,Atrib3"
   Q1 = Q1 & " ,Atrib4"
   Q1 = Q1 & " ,Atrib5"
   Q1 = Q1 & " ,Atrib6"
   Q1 = Q1 & " ,Atrib7"
   Q1 = Q1 & " ,Atrib8"
   Q1 = Q1 & " ,Atrib9"
   Q1 = Q1 & " ,Atrib10"
   Q1 = Q1 & " ,CodPlanAvanzado"
   Q1 = Q1 & " ,TipoPartida"
   Q1 = Q1 & " ,CodCtaPlanSII"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM IFRS_PlanIFRS) as Cant"
   Q1 = Q1 & " From IFRS_PlanIFRS"
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
   
   If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From IFRS_PlanIFRS"
            Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT IFRS_PlanIFRS ON "
                Q1 = Q1 & " INSERT INTO IFRS_PlanIFRS"
                Q1 = Q1 & " (idCuenta"
                Q1 = Q1 & " ,idPadre"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,Nivel"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Clasificacion"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,TipoCapPropio"
                Q1 = Q1 & " ,CodF22"
                Q1 = Q1 & " ,Atrib1"
                Q1 = Q1 & " ,Atrib2"
                Q1 = Q1 & " ,Atrib3"
                Q1 = Q1 & " ,Atrib4"
                Q1 = Q1 & " ,Atrib5"
                Q1 = Q1 & " ,Atrib6"
                Q1 = Q1 & " ,Atrib7"
                Q1 = Q1 & " ,Atrib8"
                Q1 = Q1 & " ,Atrib9"
                Q1 = Q1 & " ,Atrib10"
                Q1 = Q1 & " ,CodPlanAvanzado"
                Q1 = Q1 & " ,TipoPartida"
                Q1 = Q1 & " ,CodCtaPlanSII)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodPlanAvanzado")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaPlanSII")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT IFRS_PlanIFRS OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE IFRS_PlanIFRS"
                Q1 = Q1 & " SET idPadre = " & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,TipoCapPropio = " & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ,CodF22 = " & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ,Atrib1 = " & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ,Atrib2 = " & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ,Atrib3 = " & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ,Atrib4 = " & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ,Atrib5 = " & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ,Atrib6 = " & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ,Atrib7 = " & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ,Atrib8 = " & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ,Atrib9 = " & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ,Atrib10 = " & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,CodPlanAvanzado = '" & vFldDao(Rs("CodPlanAvanzado")) & "'"
                Q1 = Q1 & " ,TipoPartida = " & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,CodCtaPlanSII = '" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasPlanCuentasSII(DBSql As ADODB.Connection, DbAccess As Database)
    
    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM PlanCuentasSII"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From PlanCuentasSII"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT IdPlanCuentasSII"
   Q1 = Q1 & " ,CodigoSII"
   Q1 = Q1 & " ,DescripSII"
   Q1 = Q1 & " ,FmtCodigoSII"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,AnoDesde"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM PlanCuentasSII) as Cant"
   Q1 = Q1 & " FROM PlanCuentasSII"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
    If Rs.EOF = False Then
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From PlanCuentasSII"
            Q1 = Q1 & " WHERE IdPlanCuentasSII = " & vFldDao(Rs("IdPlanCuentasSII"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT PlanCuentasSII ON "
                Q1 = Q1 & " INSERT INTO PlanCuentasSII"
                Q1 = Q1 & " (IdPlanCuentasSII"
                Q1 = Q1 & " ,CodigoSII"
                Q1 = Q1 & " ,DescripSII"
                Q1 = Q1 & " ,FmtCodigoSII"
                Q1 = Q1 & " ,Clasificacion"
                Q1 = Q1 & " ,AnoDesde)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdPlanCuentasSII"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodigoSII")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("DescripSII")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("FmtCodigoSII")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("AnoDesde")) & ")"
                Q1 = Q1 & " SET IDENTITY_INSERT PlanCuentasSII OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE PlanCuentasSII"
                Q1 = Q1 & " SET CodigoSII = '" & vFldDao(Rs("CodigoSII")) & "'"
                Q1 = Q1 & " ,DescripSII = '" & vFldDao(Rs("DescripSII")) & "'"
                Q1 = Q1 & " ,FmtCodigoSII = '" & vFldDao(Rs("FmtCodigoSII")) & "'"
                Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ,AnoDesde = " & vFldDao(Rs("AnoDesde"))
                Q1 = Q1 & " WHERE IdPlanCuentasSII = " & vFldDao(Rs("IdPlanCuentasSII"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasGlosas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Glosas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Glosas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT idGlosa"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " From Glosas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Glosas"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idGlosa"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Glosas ON "
                Q1 = " INSERT INTO Glosas"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("idGlosa")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Glosas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Glosas"
                Q1 = Q1 & " SET Glosa = '" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idGlosa"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasFirmas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

   Q1 = "SELECT Patch"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,ano"
   Q1 = Q1 & " From Firmas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = '" & Ano & "'"
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Firmas"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND Tipo = '" & vFldDao(Rs("Tipo")) & "'"
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Firmas ON "
                Q1 = " INSERT INTO Firmas"
                Q1 = Q1 & " (Patch"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,ano)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("Patch")) & "'"
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("Tipo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("ano")) & "')"
                'Q1 = Q1 & " SET IDENTITY_INSERT Firmas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Firmas"
                Q1 = Q1 & " SET Patch = '" & vFldDao(Rs("Patch")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = '" & vFldDao(Rs("Ano")) & "'"
                Q1 = Q1 & " AND Tipo = '" & vFldDao(Rs("Tipo")) & "'"
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasFactorActAnual(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(IdFactorActAnual) as Cant"
    Q1 = Q1 & " From FactorActAnual"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT IdFactorActAnual"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,MesRow"
   Q1 = Q1 & " ,MesCol"
   Q1 = Q1 & " ,Factor"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM FactorActAnual) as Cant"
   Q1 = Q1 & " From FactorActAnual"
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
   
   If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From FactorActAnual"
            Q1 = Q1 & " WHERE Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND MesCol = " & vFldDao(Rs("MesCol"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                'Q1 = " SET IDENTITY_INSERT FactorActAnual ON "
                Q1 = Q1 & " INSERT INTO FactorActAnual"
                'Q1 = Q1 & " (IdFactorActAnual"
                Q1 = Q1 & " (Ano"
                Q1 = Q1 & " ,MesRow"
                Q1 = Q1 & " ,MesCol"
                Q1 = Q1 & " ,Factor)"
                Q1 = Q1 & " Values"
                'Q1 = Q1 & " (" & vFldDao(Rs("IdFactorActAnual"))
                Q1 = Q1 & " (" & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("MesRow"))
                Q1 = Q1 & " ," & vFldDao(Rs("MesCol"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Factor")))) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT FactorActAnual OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE FactorActAnual "
                Q1 = Q1 & " SET MesRow = " & vFldDao(Rs("MesRow"))
                Q1 = Q1 & " ,Factor = " & str(vFmt(vFldDao(Rs("Factor"))))
                Q1 = Q1 & " WHERE IdFactorActAnual = " & vFldDao(Rs("IdFactorActAnual"))
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND MesCol = " & vFldDao(Rs("MesCol"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasEstadoMes(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

   Q1 = "SELECT Mes"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Impreso"
   Q1 = Q1 & " ,FechaApertura"
   Q1 = Q1 & " ,FechaCierre"
   Q1 = Q1 & " ,FechaImpresion"
   Q1 = Q1 & " From EstadoMes"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From EstadoMes"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT EstadoMes ON "
                Q1 = " INSERT INTO EstadoMes"
                Q1 = Q1 & " (Mes"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Impreso"
                Q1 = Q1 & " ,FechaApertura"
                Q1 = Q1 & " ,FechaCierre"
                Q1 = Q1 & " ,FechaImpresion)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("Mes"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Impreso"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaApertura"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaCierre"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaImpresion")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT EstadoMes OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE EstadoMes"
                Q1 = Q1 & " SET Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Impreso = " & vFldDao(Rs("Impreso"))
                Q1 = Q1 & " ,FechaApertura = " & vFldDao(Rs("FechaApertura"))
                Q1 = Q1 & " ,FechaCierre = " & vFldDao(Rs("FechaCierre"))
                Q1 = Q1 & " ,FechaImpresion = " & vFldDao(Rs("FechaImpresion"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasEquivalencia(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(idMoneda) as Cant"
    Q1 = Q1 & " From Equivalencia"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)


   Q1 = "SELECT idMoneda"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM Equivalencia) as Cant"
   Q1 = Q1 & " From Equivalencia"
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
   
   If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Equivalencia"
            Q1 = Q1 & " WHERE idMoneda = " & vFldDao(Rs("idMoneda"))
            Q1 = Q1 & " AND Fecha = " & vFldDao(Rs("Fecha"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Equivalencia ON "
                Q1 = " INSERT INTO Equivalencia"
                Q1 = Q1 & " (idMoneda"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,Valor)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idMoneda"))
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Valor")))) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Equivalencia OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Equivalencia"
                Q1 = Q1 & " SET Valor = " & str(vFmt(vFldDao(Rs("Valor"))))
                Q1 = Q1 & " WHERE idMoneda = " & vFldDao(Rs("idMoneda"))
                Q1 = Q1 & " AND Fecha = " & vFldDao(Rs("Fecha"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasEquipos()

'   Q1 = "SELECT PC"
'   Q1 = Q1 & " ,MAC"
'   Q1 = Q1 & " ,CodPC"
'   Q1 = Q1 & " ,Aut"
'   Q1 = Q1 & " From Equipos"
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT Equipos ON "
'      Q1 = Q1 & " INSERT INTO Equipos"
'      Q1 = Q1 & " (PC"
'      Q1 = Q1 & " ,MAC"
'      Q1 = Q1 & " ,CodPC"
'      Q1 = Q1 & " ,Aut)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " ('" & vFld(Rs("PC")) & "'"
'      Q1 = Q1 & " ,'" & vFld(Rs("MAC")) & "'"
'      Q1 = Q1 & " ,'" & vFld(Rs("CodPC")) & "'"
'      Q1 = Q1 & " ," & vFld(Rs("Aut")) & ")"
'      Q1 = Q1 & " SET IDENTITY_INSERT Equipos OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)


End Sub

Public Sub TrasDocCuotas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'DocCuotas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE DocCuotas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdDocCuota"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdDoc"
   Q1 = Q1 & " ,NumCuota"
   Q1 = Q1 & " ,FechaExigPago"
   Q1 = Q1 & " ,MontoCuota"
   Q1 = Q1 & " ,FechaIngPercibido"
   Q1 = Q1 & " ,IdCompPago"
   Q1 = Q1 & " ,IdLibCaja"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " From DocCuotas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From DocCuotas"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDocCuota"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT DocCuotas ON "
                Q1 = " INSERT INTO DocCuotas"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdDoc"
                Q1 = Q1 & " ,NumCuota"
                Q1 = Q1 & " ,FechaExigPago"
                Q1 = Q1 & " ,MontoCuota"
                Q1 = Q1 & " ,FechaIngPercibido"
                Q1 = Q1 & " ,IdCompPago"
                Q1 = Q1 & " ,IdLibCaja"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumCuota"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaExigPago"))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoCuota"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaIngPercibido"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdLibCaja"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDocCuota")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT DocCuotas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE DocCuotas"
                Q1 = Q1 & " SET IdDoc = " & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ,NumCuota = " & vFldDao(Rs("NumCuota"))
                Q1 = Q1 & " ,FechaExigPago = " & vFldDao(Rs("FechaExigPago"))
                Q1 = Q1 & " ,MontoCuota = " & vFldDao(Rs("MontoCuota"))
                Q1 = Q1 & " ,FechaIngPercibido = " & vFldDao(Rs("FechaIngPercibido"))
                Q1 = Q1 & " ,IdCompPago = " & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ,IdLibCaja = " & vFldDao(Rs("IdLibCaja"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDocCuota"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE DC"
    Q1 = Q1 & " SET    DC.IdDoc = ISNULL(D.IdDoc,DC.IdDoc),"
    Q1 = Q1 & "        DC.IdCompPago = ISNULL(C.IdComp,DC.IdCompPago),"
    Q1 = Q1 & "        DC.IdLibCaja = ISNULL(l.IdLibroCaja, DC.IdLibCaja)"
    Q1 = Q1 & " FROM DocCuotas DC"
    Q1 = Q1 & " LEFT JOIN Documento D ON D.IdTras = DC.IdDoc AND D.IdEmpresa = DC.IdEmpresa AND D.Ano = DC.Ano"
    Q1 = Q1 & " LEFT JOIN Comprobante C ON C.IdTras = DC.IdCompPago AND C.IdEmpresa = DC.IdEmpresa AND C.Ano = DC.Ano"
    Q1 = Q1 & " LEFT JOIN LibroCaja L ON L.IdTras = DC.IdLibCaja AND L.IdEmpresa = DC.IdEmpresa AND L.Ano = DC.Ano"
    Q1 = Q1 & " WHERE DC.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND DC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasDetSaldosAp(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'DetSaldosAp' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE DetSaldosAp ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT Id"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,IdEntidad"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,Saldo"
   Q1 = Q1 & " FROM DetSaldosAp"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From DetSaldosAp"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("Id"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT DetSaldosAp ON "
                Q1 = " INSERT INTO DetSaldosAp"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,IdEntidad"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,Saldo"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ," & vFldDao(Rs("Saldo"))
                Q1 = Q1 & " ," & vFldDao(Rs("Id")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT DetSaldosAp OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE DetSaldosAp"
                Q1 = Q1 & " SET IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,IdEntidad = " & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,Saldo = " & vFldDao(Rs("Saldo"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("Id"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE DS"
    Q1 = Q1 & " SET    DS.IdCuenta = ISNULL(CU.idCuenta,DS.IdCuenta),"
    Q1 = Q1 & "        DS.IdEntidad = ISNULL(e.IdEntidad, DS.IdEntidad)"
    Q1 = Q1 & " FROM DetSaldosAp DS"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = DS.IdCuenta AND CU.IdEmpresa = DS.IdEmpresa AND CU.Ano = DS.Ano"
    Q1 = Q1 & " LEFT JOIN Entidades E ON E.IdTras = DS.IdEntidad AND E.IdEmpresa = DS.IdEmpresa"
    Q1 = Q1 & " WHERE DS.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   DS.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasDetPercepciones(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = "DELETE FROM DetPercepciones WHERE IDPerc IN (SELECT IDPerc FROM Percepciones WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano & ") "
    Call ExecSQL(DBSql, Q1)
   
   Q1 = "SELECT IDPerc"
   Q1 = Q1 & " ,CodDet"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " From DetPercepciones"
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
                      
                'Q1 = " SET IDENTITY_INSERT DetPercepciones ON "
                Q1 = " INSERT INTO DetPercepciones"
                Q1 = Q1 & " (IDPerc"
                Q1 = Q1 & " ,CodDet"
                Q1 = Q1 & " ,Valor)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IDPerc"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodDet"))
                Q1 = Q1 & " ," & vFldDao(Rs("Valor")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT DetPercepciones OFF  "
                Call ExecSQL(DBSql, Q1)

      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   
    Q1 = "UPDATE DP"
    Q1 = Q1 & " SET DP.IdPerc = IsNull(P.IdPerc, DP.IdPerc)"
    Q1 = Q1 & " FROM DetPercepciones DP"
    Q1 = Q1 & " LEFT JOIN Percepciones P ON P.IdTras = DP.IDPerc AND P.IdEmpresa = " & IdEmpresa & " AND P.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasDetCartola(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'DetCartola' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE DetCartola ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdDetCartola"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,IdCartola"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,Detalle"
   Q1 = Q1 & " ,NumDoc"
   Q1 = Q1 & " ,Cargo"
   Q1 = Q1 & " ,Abono"
   Q1 = Q1 & " ,IdMov"
   Q1 = Q1 & " From DetCartola"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From DetCartola"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDetCartola"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT DetCartola ON "
                Q1 = " INSERT INTO DetCartola"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdCartola"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,Detalle"
                Q1 = Q1 & " ,NumDoc"
                Q1 = Q1 & " ,Cargo"
                Q1 = Q1 & " ,Abono"
                Q1 = Q1 & " ,IdMov"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCartola"))
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Detalle")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("NumDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("Cargo"))
                Q1 = Q1 & " ," & vFldDao(Rs("Abono"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdMov"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDetCartola")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT DetCartola OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE DetCartola"
                Q1 = Q1 & " SET IdCartola = " & vFldDao(Rs("IdCartola"))
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Detalle = '" & vFldDao(Rs("Detalle")) & "'"
                Q1 = Q1 & " ,NumDoc = " & vFldDao(Rs("NumDoc"))
                Q1 = Q1 & " ,Cargo = " & vFldDao(Rs("Cargo"))
                Q1 = Q1 & " ,Abono = " & vFldDao(Rs("Abono"))
                Q1 = Q1 & " ,IdMov = " & vFldDao(Rs("IdMov"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDetCartola"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE DC"
    Q1 = Q1 & " SET    DC.IdCartola = C.IdCartola,"
    Q1 = Q1 & "        DC.idMov = MC.idMov"
    Q1 = Q1 & " FROM DetCartola DC"
    Q1 = Q1 & " LEFT JOIN Cartola C ON C.IdTras = DC.IdCartola AND C.IdEmpresa = DC.IdEmpresa AND C.Ano = DC.Ano"
    Q1 = Q1 & " LEFT JOIN MovComprobante MC ON MC.IdTras = DC.IdMov AND MC.IdEmpresa = DC.IdEmpresa AND MC.Ano = DC.Ano"
    Q1 = Q1 & " WHERE DC.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   DC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasDetCapPropioSimpl(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'DetCapPropioSimpl' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE DetCapPropioSimpl ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdDetCapPropioSimpl"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,TipoDetCPS"
   Q1 = Q1 & " ,IngresoManual"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,IdMovComp"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " ,Descrip"
   Q1 = Q1 & " From DetCapPropioSimpl"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From DetCapPropioSimpl"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDetCapPropioSimpl"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT DetCapPropioSimpl ON "
                Q1 = " INSERT INTO DetCapPropioSimpl"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,TipoDetCPS"
                Q1 = Q1 & " ,IngresoManual"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,CodCuenta"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,IdMovComp"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDetCPS"))
                Q1 = Q1 & " ," & vFldDao(Rs("IngresoManual"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdMovComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Valor"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdDetCapPropioSimpl")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT DetCapPropioSimpl OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE DetCapPropioSimpl"
                Q1 = Q1 & " SET TipoDetCPS = " & vFldDao(Rs("TipoDetCPS"))
                Q1 = Q1 & " ,IngresoManual = " & vFldDao(Rs("IngresoManual"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,IdMovComp = " & vFldDao(Rs("IdMovComp"))
                Q1 = Q1 & " ,Valor = " & vFldDao(Rs("Valor"))
                Q1 = Q1 & " ,Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDetCapPropioSimpl"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE DC"
    Q1 = Q1 & " SET DC.IdCuenta = C.idCuenta,"
    Q1 = Q1 & "        DC.IdMovComp = MC.idMov"
    Q1 = Q1 & " FROM DetCapPropioSimpl DC"
    Q1 = Q1 & " LEFT JOIN Cuentas C ON C.IdTras = DC.IdCuenta AND C.IdEmpresa = DC.IdEmpresa AND C.Ano = DC.Ano"
    Q1 = Q1 & " LEFT JOIN MovComprobante MC ON MC.IdTras = DC.IdMovComp AND MC.IdEmpresa = DC.IdEmpresa"
    Q1 = Q1 & " WHERE DC.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND DC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCuentasRazon(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

   Q1 = "SELECT IdRazon"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,NumDenom"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " ,Operador"
   Q1 = Q1 & " From CuentasRazon"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CuentasRazon"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdRazon = " & vFldDao(Rs("IdRazon"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CuentasRazon ON "
                Q1 = Q1 & " INSERT INTO CuentasRazon"
                Q1 = Q1 & " (IdRazon"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,NumDenom"
                Q1 = Q1 & " ,CodCuenta"
                Q1 = Q1 & " ,Operador)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdRazon"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("NumDenom"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Operador")) & "')"
                'Q1 = Q1 & " SET IDENTITY_INSERT CuentasRazon OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CuentasRazon"
                Q1 = Q1 & " SET NumDenom = " & vFldDao(Rs("NumDenom"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ,Operador = '" & vFldDao(Rs("Operador")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdRazon = " & vFldDao(Rs("IdRazon"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasCuentasBasicas(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CuentasBasicas' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CuentasBasicas ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT Id"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,TipoLib"
   Q1 = Q1 & " ,TipoValor"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,IdCuentaOld"
   Q1 = Q1 & " From CuentasBasicas"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CuentasBasicas"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("Id"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CuentasBasicas ON "
                Q1 = Q1 & " INSERT INTO CuentasBasicas"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,TipoValor"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,IdCuentaOld"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoValor"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("Id")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CuentasBasicas OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CuentasBasicas"
                Q1 = Q1 & " SET Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,TipoValor = " & vFldDao(Rs("TipoValor"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,IdCuentaOld = " & vFldDao(Rs("IdCuentaOld"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("Id"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE CB"
    Q1 = Q1 & " SET CB.idCuenta = ISNULL(C.idCuenta,CB.idCuenta)"
    Q1 = Q1 & " FROM CuentasBasicas CB"
    Q1 = Q1 & " INNER JOIN Cuentas C ON C.IdTras = CB.IdCuenta AND C.IdEmpresa = CB.IdEmpresa AND C.Ano = CB.Ano"
    Q1 = Q1 & " WHERE CB.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND CB.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)

    Q1 = "UPDATE CB"
    Q1 = Q1 & " SET    CB.IdCuentaOld = ISNULL(CU.IdCuenta, Cb.IdCuentaOld)"
    Q1 = Q1 & " FROM CuentasBasicas CB"
    Q1 = Q1 & " INNER JOIN Cuentas CU ON CU.IdTras = CB.IdCuentaOld AND CU.IdEmpresa = CB.IdEmpresa AND CU.Ano = CB.Ano"
    Q1 = Q1 & " WHERE Cb.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND CB.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCtasAjustesExContRLI(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CtasAjustesExContRLI' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CtasAjustesExContRLI ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdCtaAjustesRLI"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,TipoAjuste"
   Q1 = Q1 & " ,IdGrupo"
   Q1 = Q1 & " ,IdItem"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " From CtasAjustesExContRLI"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CtasAjustesExContRLI"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCtaAjustesRLI"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CtasAjustesExContRLI ON "
                Q1 = " INSERT INTO CtasAjustesExContRLI"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,TipoAjuste"
                Q1 = Q1 & " ,IdGrupo"
                Q1 = Q1 & " ,IdItem"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,CodCuenta"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdItem"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdCtaAjustesRLI")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CtasAjustesExContRLI OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CtasAjustesExContRLI"
                Q1 = Q1 & " SET TipoAjuste = " & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ,IdGrupo = " & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ,IdItem = " & vFldDao(Rs("IdItem"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCtaAjustesRLI"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE CA"
    Q1 = Q1 & " SET CA.IdCuenta = C.IdCuenta"
    Q1 = Q1 & " FROM CtasAjustesExContRLI CA"
    Q1 = Q1 & " LEFT JOIN Cuentas C ON C.IdTras = CA.IdCuenta AND C.IdEmpresa = CA.IdEmpresa AND C.Ano = CA.Ano"
    Q1 = Q1 & " WHERE CA.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND CA.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCtasAjustesExCont(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CtasAjustesExCont' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CtasAjustesExCont ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdCtaAjustes"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,TipoAjuste"
   Q1 = Q1 & " ,IdItem"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " From CtasAjustesExCont"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CtasAjustesExCont"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCtaAjustes"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
              'Q1 = " SET IDENTITY_INSERT CtasAjustesExCont ON "
              Q1 = " INSERT INTO CtasAjustesExCont"
              Q1 = Q1 & " (IdEmpresa"
              Q1 = Q1 & " ,Ano"
              Q1 = Q1 & " ,TipoAjuste"
              Q1 = Q1 & " ,IdItem"
              Q1 = Q1 & " ,IdCuenta"
              Q1 = Q1 & " ,CodCuenta"
              Q1 = Q1 & " ,IdTras)"
              Q1 = Q1 & " Values"
              Q1 = Q1 & " (" & IdEmpresa
              Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
              Q1 = Q1 & " ," & vFldDao(Rs("TipoAjuste"))
              Q1 = Q1 & " ," & vFldDao(Rs("IdItem"))
              Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
              Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
              Q1 = Q1 & " ," & vFldDao(Rs("IdCtaAjustes")) & ")"
              'Q1 = Q1 & " SET IDENTITY_INSERT CtasAjustesExCont OFF  "
              Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CtasAjustesExCont"
                Q1 = Q1 & " SET TipoAjuste = " & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ,IdItem = " & vFldDao(Rs("IdItem"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCtaAjustes"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE CA"
    Q1 = Q1 & " SET CA.IdCuenta = C.IdCuenta"
    Q1 = Q1 & " FROM CtasAjustesExCont CA"
    Q1 = Q1 & " LEFT JOIN Cuentas C ON C.IdTras = CA.IdCuenta AND C.IdEmpresa = CA.IdEmpresa AND C.Ano = CA.Ano"
    Q1 = Q1 & " WHERE CA.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND CA.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCT_MovComprobanteBase()

'   Q1 = "SELECT IdMov"
'   Q1 = Q1 & " ,IdEmpresa"
'   Q1 = Q1 & " ,IdComp"
'   Q1 = Q1 & " ,Orden"
'   Q1 = Q1 & " ,IdCuenta"
'   Q1 = Q1 & " ,CodCuenta"
'   Q1 = Q1 & " ,Debe"
'   Q1 = Q1 & " ,Haber"
'   Q1 = Q1 & " ,Glosa"
'   Q1 = Q1 & " ,IdCCosto"
'   Q1 = Q1 & " ,IdAreaNeg"
'   Q1 = Q1 & " ,Conciliado"
'   Q1 = Q1 & " From CT_MovComprobanteBase"
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT CT_MovComprobanteBase ON "
'      Q1 = Q1 & " INSERT INTO CT_MovComprobanteBase"
'      Q1 = Q1 & " (IdEmpresa"
'      Q1 = Q1 & " ,IdComp"
'      Q1 = Q1 & " ,Orden"
'      Q1 = Q1 & " ,IdCuenta"
'      Q1 = Q1 & " ,CodCuenta"
'      Q1 = Q1 & " ,Debe"
'      Q1 = Q1 & " ,Haber"
'      Q1 = Q1 & " ,Glosa"
'      Q1 = Q1 & " ,IdCCosto"
'      Q1 = Q1 & " ,IdAreaNeg"
'      Q1 = Q1 & " ,Conciliado)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("IdEmpresa"))
'      Q1 = Q1 & " ," & vFld(Rs("IdComp"))
'      Q1 = Q1 & " ," & vFld(Rs("Orden"))
'      Q1 = Q1 & " ," & vFld(Rs("IdCuenta"))
'      Q1 = Q1 & " ,'" & vFld(Rs("CodCuenta")) & "'"
'      Q1 = Q1 & " ," & vFld(Rs("Debe"))
'      Q1 = Q1 & " ," & vFld(Rs("Haber"))
'      Q1 = Q1 & " ,'" & vFld(Rs("Glosa")) & "'"
'      Q1 = Q1 & " ," & vFld(Rs("IdCCosto"))
'      Q1 = Q1 & " ," & vFld(Rs("IdAreaNeg"))
'      Q1 = Q1 & " ," & vFld(Rs("Conciliado")) & ")"
'      Q1 = Q1 & " SET IDENTITY_INSERT CT_MovComprobanteBase OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)


End Sub

Public Sub TrasCT_MovComprobante(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CT_MovComprobante' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CT_MovComprobante ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdMov"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,IdComp"
   Q1 = Q1 & " ,Orden"
   Q1 = Q1 & " ,IdCuenta"
   Q1 = Q1 & " ,CodCuenta"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,IdCCosto"
   Q1 = Q1 & " ,IdAreaNeg"
   Q1 = Q1 & " ,Conciliado"
   Q1 = Q1 & " From CT_MovComprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CT_MovComprobante"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMov"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CT_MovComprobante ON "
                Q1 = " INSERT INTO CT_MovComprobante"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,IdComp"
                Q1 = Q1 & " ,Orden"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,CodCuenta"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,IdCCosto"
                Q1 = Q1 & " ,IdAreaNeg"
                Q1 = Q1 & " ,Conciliado"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdCCosto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAreaNeg"))
                Q1 = Q1 & " ," & vFldDao(Rs("Conciliado"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdMov")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CT_MovComprobante OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CT_MovComprobante"
                Q1 = Q1 & " SET IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,Orden = " & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,CodCuenta = '" & vFldDao(Rs("CodCuenta")) & "'"
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,Glosa = '" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ,IdCCosto = " & vFldDao(Rs("IdCCosto"))
                Q1 = Q1 & " ,IdAreaNeg = " & vFldDao(Rs("IdAreaNeg"))
                Q1 = Q1 & " ,Conciliado = " & vFldDao(Rs("Conciliado"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMov"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE CT"
    Q1 = Q1 & " SET    CT.IdComp = ISNULL(CTC.IdComp, CT.IdComp),"
    Q1 = Q1 & "        CT.IdCuenta = ISNULL(CU.idCuenta, CT.IdCuenta),"
    Q1 = Q1 & "        CT.IdCCosto = ISNULL(CC.IdCCosto, CT.IdCCosto),"
    Q1 = Q1 & "        CT.IdAreaNeg = IsNull(cc.IdCCosto, CT.IdAreaNeg)"
    Q1 = Q1 & " FROM ((((CT_MovComprobante CT"
    Q1 = Q1 & " LEFT JOIN CT_Comprobante CTC ON CTC.IdTras = CT.IdComp)"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = CT.IdCuenta)"
    Q1 = Q1 & " LEFT JOIN CentroCosto CC ON CC.IdTras = CT.IdCCosto)"
    Q1 = Q1 & " LEFT JOIN AreaNegocio AN ON AN.IdTras = CT.IdAreaNeg)"
    Q1 = Q1 & " WHERE CT.IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)

End Sub

Public Sub TrasCT_ComprobanteBase(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CT_ComprobanteBase' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CT_ComprobanteBase ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdComp"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Correlativo"
    Q1 = Q1 & " ,Nombre"
    Q1 = Q1 & " ,Descrip"
    Q1 = Q1 & " ,Fecha"
    Q1 = Q1 & " ,Tipo"
    Q1 = Q1 & " ,Estado"
    Q1 = Q1 & " ,Glosa"
    Q1 = Q1 & " ,TotalDebe"
    Q1 = Q1 & " ,TotalHaber"
    Q1 = Q1 & " ,IdUsuario"
    Q1 = Q1 & " From CT_ComprobanteBase"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CT_ComprobanteBase"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CT_ComprobanteBase ON "
                Q1 = " INSERT INTO CT_ComprobanteBase"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Correlativo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,TotalDebe"
                Q1 = Q1 & " ,TotalHaber"
                Q1 = Q1 & " ,IdUsuario)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Correlativo"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TotalDebe"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotalHaber"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CT_ComprobanteBase OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CT_ComprobanteBase "
                Q1 = Q1 & " SET Correlativo = " & vFldDao(Rs("Correlativo"))
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Glosa = '" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ,TotalDebe = " & vFldDao(Rs("TotalDebe"))
                Q1 = Q1 & " ,TotalHaber = " & vFldDao(Rs("TotalHaber"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasCT_Comprobante(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
   Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CT_Comprobante' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CT_Comprobante ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdComp"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Correlativo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descrip"
   Q1 = Q1 & " ,Fecha"
   Q1 = Q1 & " ,Tipo"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Glosa"
   Q1 = Q1 & " ,TotalDebe"
   Q1 = Q1 & " ,TotalHaber"
   Q1 = Q1 & " ,IdUsuario"
   Q1 = Q1 & " ,IdCompOld"
   Q1 = Q1 & " From CT_Comprobante"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CT_Comprobante"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CT_Comprobante ON "
                Q1 = " INSERT INTO CT_Comprobante"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Correlativo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Fecha"
                Q1 = Q1 & " ,Tipo"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,TotalDebe"
                Q1 = Q1 & " ,TotalHaber"
                Q1 = Q1 & " ,IdUsuario"
                Q1 = Q1 & " ,IdCompOld"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Correlativo"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TotalDebe"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotalHaber"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CT_Comprobante OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CT_Comprobante"
                Q1 = Q1 & " SET Correlativo = " & vFldDao(Rs("Correlativo"))
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descrip = '" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Glosa = '" & vFldDao(Rs("Glosa")) & "'"
                Q1 = Q1 & " ,TotalDebe = " & vFldDao(Rs("TotalDebe"))
                Q1 = Q1 & " ,TotalHaber = " & vFldDao(Rs("TotalHaber"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,IdCompOld = " & vFldDao(Rs("IdCompOld"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE CT"
    Q1 = Q1 & " Set CT.IdCompOld = IsNull(CTC.IdCompOld, CT.IdCompOld)"
    Q1 = Q1 & " FROM (CT_Comprobante CT"
    Q1 = Q1 & " LEFT JOIN CT_Comprobante CTC ON CTC.IdCompOld = CT.IdTras)"
    Q1 = Q1 & " WHERE CT.IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasControlEmpresa(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(IdEmpresa) as Cant"
    Q1 = Q1 & " From ControlEmpresa"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

    Q1 = "SELECT IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,RazonSocial"
    Q1 = Q1 & " ,RUT"
    Q1 = Q1 & " ,Mes1"
    Q1 = Q1 & " ,Mes2"
    Q1 = Q1 & " ,Mes3"
    Q1 = Q1 & " ,Mes4"
    Q1 = Q1 & " ,Mes5"
    Q1 = Q1 & " ,Mes6"
    Q1 = Q1 & " ,Mes7"
    Q1 = Q1 & " ,Mes8"
    Q1 = Q1 & " ,Mes9"
    Q1 = Q1 & " ,Mes10"
    Q1 = Q1 & " ,Mes11"
    Q1 = Q1 & " ,Mes12"
    Q1 = Q1 & " ,AF_Depreciacion"
    Q1 = Q1 & " ,AF_CM"
    Q1 = Q1 & " ,AF_33BisLir"
    Q1 = Q1 & " ,CM_Activos"
    Q1 = Q1 & " ,CM_Pasivos"
    Q1 = Q1 & " ,BalDefinitivo"
    Q1 = Q1 & " ,CPT_Municip"
    Q1 = Q1 & " ,F22Renta"
    Q1 = Q1 & " ,AjustesIFRS"
    Q1 = Q1 & " ,CalcPropIVA"
    Q1 = Q1 & " ,(SELECT COUNT(*) FROM ControlEmpresa) as Cant"
    Q1 = Q1 & " From ControlEmpresa"
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
   
   If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From ControlEmpresa"
            Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
              'Q1 = " SET IDENTITY_INSERT ControlEmpresa ON "
              Q1 = " INSERT INTO ControlEmpresa"
              Q1 = Q1 & " (IdEmpresa"
              Q1 = Q1 & " ,Ano"
              Q1 = Q1 & " ,RazonSocial"
              Q1 = Q1 & " ,RUT"
              Q1 = Q1 & " ,Mes1"
              Q1 = Q1 & " ,Mes2"
              Q1 = Q1 & " ,Mes3"
              Q1 = Q1 & " ,Mes4"
              Q1 = Q1 & " ,Mes5"
              Q1 = Q1 & " ,Mes6"
              Q1 = Q1 & " ,Mes7"
              Q1 = Q1 & " ,Mes8"
              Q1 = Q1 & " ,Mes9"
              Q1 = Q1 & " ,Mes10"
              Q1 = Q1 & " ,Mes11"
              Q1 = Q1 & " ,Mes12"
              Q1 = Q1 & " ,AF_Depreciacion"
              Q1 = Q1 & " ,AF_CM"
              Q1 = Q1 & " ,AF_33BisLir"
              Q1 = Q1 & " ,CM_Activos"
              Q1 = Q1 & " ,CM_Pasivos"
              Q1 = Q1 & " ,BalDefinitivo"
              Q1 = Q1 & " ,CPT_Municip"
              Q1 = Q1 & " ,F22Renta"
              Q1 = Q1 & " ,AjustesIFRS"
              Q1 = Q1 & " ,CalcPropIVA)"
              Q1 = Q1 & " Values"
              Q1 = Q1 & " (" & gEmpresa.id
              Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
              Q1 = Q1 & " ,'" & vFldDao(Rs("RazonSocial")) & "'"
              Q1 = Q1 & " ,'" & vFldDao(Rs("RUT")) & "'"
              Q1 = Q1 & " ," & vFldDao(Rs("Mes1"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes2"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes3"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes4"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes5"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes6"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes7"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes8"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes9"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes10"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes11"))
              Q1 = Q1 & " ," & vFldDao(Rs("Mes12"))
              Q1 = Q1 & " ," & vFldDao(Rs("AF_Depreciacion"))
              Q1 = Q1 & " ," & vFldDao(Rs("AF_CM"))
              Q1 = Q1 & " ," & vFldDao(Rs("AF_33BisLir"))
              Q1 = Q1 & " ," & vFldDao(Rs("CM_Activos"))
              Q1 = Q1 & " ," & vFldDao(Rs("CM_Pasivos"))
              Q1 = Q1 & " ," & vFldDao(Rs("BalDefinitivo"))
              Q1 = Q1 & " ," & vFldDao(Rs("CPT_Municip"))
              Q1 = Q1 & " ," & vFldDao(Rs("F22Renta"))
              Q1 = Q1 & " ," & vFldDao(Rs("AjustesIFRS"))
              Q1 = Q1 & " ," & vFldDao(Rs("CalcPropIVA")) & ")"
              'Q1 = Q1 & " SET IDENTITY_INSERT ControlEmpresa OFF  "
              Call ExecSQL(DBSql, Q1)
                
            Else
            
              Q1 = " UPDATE ControlEmpresa"
              Q1 = Q1 & " SET RazonSocial = '" & vFldDao(Rs("RazonSocial")) & "'"
              Q1 = Q1 & " , RUT = '" & vFldDao(Rs("RUT")) & "'"
              Q1 = Q1 & " , Mes1 = " & vFldDao(Rs("Mes1"))
              Q1 = Q1 & " , Mes2 = " & vFldDao(Rs("Mes2"))
              Q1 = Q1 & " , Mes3 = " & vFldDao(Rs("Mes3"))
              Q1 = Q1 & " , Mes4 = " & vFldDao(Rs("Mes4"))
              Q1 = Q1 & " , Mes5 = " & vFldDao(Rs("Mes5"))
              Q1 = Q1 & " , Mes6 = " & vFldDao(Rs("Mes6"))
              Q1 = Q1 & " , Mes7 = " & vFldDao(Rs("Mes7"))
              Q1 = Q1 & " , Mes8 = " & vFldDao(Rs("Mes8"))
              Q1 = Q1 & " , Mes9 = " & vFldDao(Rs("Mes9"))
              Q1 = Q1 & " , Mes10 = " & vFldDao(Rs("Mes10"))
              Q1 = Q1 & " , Mes11 = " & vFldDao(Rs("Mes11"))
              Q1 = Q1 & " , Mes12 = " & vFldDao(Rs("Mes12"))
              Q1 = Q1 & " , AF_Depreciacion = " & vFldDao(Rs("AF_Depreciacion"))
              Q1 = Q1 & " , AF_CM = " & vFldDao(Rs("AF_CM"))
              Q1 = Q1 & " , AF_33BisLir = " & vFldDao(Rs("AF_33BisLir"))
              Q1 = Q1 & " , CM_Activos =" & vFldDao(Rs("CM_Activos"))
              Q1 = Q1 & " , CM_Pasivos = " & vFldDao(Rs("CM_Pasivos"))
              Q1 = Q1 & " , BalDefinitivo = " & vFldDao(Rs("BalDefinitivo"))
              Q1 = Q1 & " , CPT_Municip = " & vFldDao(Rs("CPT_Municip"))
              Q1 = Q1 & " , F22Renta = " & vFldDao(Rs("F22Renta"))
              Q1 = Q1 & " , AjustesIFRS =" & vFldDao(Rs("AjustesIFRS"))
              Q1 = Q1 & " , CalcPropIVA =" & vFldDao(Rs("CalcPropIVA"))
              Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
              Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
              Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasContactos(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Contactos' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Contactos ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT idContacto"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,idEntidad"
    Q1 = Q1 & " ,Nombre"
    Q1 = Q1 & " ,Telefono"
    Q1 = Q1 & " ,Cargo"
    Q1 = Q1 & " From Contactos"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Contactos"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idContacto"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Contactos ON "
                Q1 = " INSERT INTO Contactos"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,idEntidad"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Telefono"
                Q1 = Q1 & " ,Cargo"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("idEntidad"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Telefono")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Cargo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("idContacto")) & "')"
                'Q1 = Q1 & " SET IDENTITY_INSERT Contactos OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Contactos"
                Q1 = Q1 & " SET idEntidad = " & vFldDao(Rs("idEntidad"))
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Telefono = '" & vFldDao(Rs("Telefono")) & "'"
                Q1 = Q1 & " ,Cargo = '" & vFldDao(Rs("Cargo")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("idContacto"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE C"
    Q1 = Q1 & " SET C.IdEntidad = E.IdEntidad"
    Q1 = Q1 & " FROM Contactos C"
    Q1 = Q1 & " LEFT JOIN Entidades E ON E.IdTras = C.idEntidad AND E.IdEmpresa = C.IdEmpresa"
    Q1 = Q1 & " WHERE c.IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasComprobante(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Comprobante' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Comprobante ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdComp"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,Correlativo"
    Q1 = Q1 & " ,Fecha"
    Q1 = Q1 & " ,Tipo"
    Q1 = Q1 & " ,Estado"
    Q1 = Q1 & " ,Glosa"
    Q1 = Q1 & " ,TotalDebe"
    Q1 = Q1 & " ,TotalHaber"
    Q1 = Q1 & " ,IdUsuario"
    Q1 = Q1 & " ,FechaCreacion"
    Q1 = Q1 & " ,ImpResumido"
    Q1 = Q1 & " ,EsCCMM"
    Q1 = Q1 & " ,FechaImport"
    Q1 = Q1 & " ,TipoAjuste"
    Q1 = Q1 & " ,OtrosIngEg14TER"
    Q1 = Q1 & " From Comprobante"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Comprobante"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
              'Q1 = " SET IDENTITY_INSERT Comprobante ON "
              Q1 = " INSERT INTO Comprobante"
              Q1 = Q1 & " (IdEmpresa"
              Q1 = Q1 & " ,Ano"
              Q1 = Q1 & " ,Correlativo"
              Q1 = Q1 & " ,Fecha"
              Q1 = Q1 & " ,Tipo"
              Q1 = Q1 & " ,Estado"
              Q1 = Q1 & " ,Glosa"
              Q1 = Q1 & " ,TotalDebe"
              Q1 = Q1 & " ,TotalHaber"
              Q1 = Q1 & " ,IdUsuario"
              Q1 = Q1 & " ,FechaCreacion"
              Q1 = Q1 & " ,ImpResumido"
              Q1 = Q1 & " ,EsCCMM"
              Q1 = Q1 & " ,FechaImport"
              Q1 = Q1 & " ,TipoAjuste"
              Q1 = Q1 & " ,OtrosIngEg14TER"
              Q1 = Q1 & " ,IdTras)"
              Q1 = Q1 & " Values"
              Q1 = Q1 & " (" & IdEmpresa
              Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
              Q1 = Q1 & " ," & vFldDao(Rs("Correlativo"))
              Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
              Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
              Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
              Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
              Q1 = Q1 & " ," & vFldDao(Rs("TotalDebe"))
              Q1 = Q1 & " ," & vFldDao(Rs("TotalHaber"))
              Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario"))
              Q1 = Q1 & " ," & vFldDao(Rs("FechaCreacion"))
              Q1 = Q1 & " ," & vFldDao(Rs("ImpResumido"))
              Q1 = Q1 & " ," & vFldDao(Rs("EsCCMM"))
              Q1 = Q1 & " ," & vFldDao(Rs("FechaImport"))
              Q1 = Q1 & " ," & vFldDao(Rs("TipoAjuste"))
              Q1 = Q1 & " ," & vFldDao(Rs("OtrosIngEg14TER"))
              Q1 = Q1 & " ," & vFldDao(Rs("IdComp")) & ")"
              'Q1 = Q1 & " SET IDENTITY_INSERT Comprobante OFF  "
              Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Comprobante"
                Q1 = Q1 & " SET Correlativo = " & vFldDao(Rs("Correlativo"))
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Glosa = '" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
                Q1 = Q1 & " ,TotalDebe = " & vFldDao(Rs("TotalDebe"))
                Q1 = Q1 & " ,TotalHaber = " & vFldDao(Rs("TotalHaber"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,FechaCreacion = " & vFldDao(Rs("FechaCreacion"))
                Q1 = Q1 & " ,ImpResumido = " & vFldDao(Rs("ImpResumido"))
                Q1 = Q1 & " ,EsCCMM = " & vFldDao(Rs("EsCCMM"))
                Q1 = Q1 & " ,FechaImport = " & vFldDao(Rs("FechaImport"))
                Q1 = Q1 & " ,TipoAjuste = " & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ,OtrosIngEg14TER = " & vFldDao(Rs("OtrosIngEg14TER"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasColores(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = "SELECT Nivel"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Color"
    Q1 = Q1 & " From Colores"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Colores"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Nivel = " & vFldDao(Rs("Nivel"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Colores ON "
                Q1 = " INSERT INTO Colores"
                Q1 = Q1 & " (Nivel"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Color)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Color")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Colores OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Colores"
                Q1 = Q1 & " SET Color = " & vFldDao(Rs("Color"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Nivel = " & vFldDao(Rs("Nivel"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub



Public Sub TrasCodActiv(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(Codigo) as Cant"
    Q1 = Q1 & " From CodActiv"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)
    
    Q1 = "SELECT Codigo"
    Q1 = Q1 & " ,Descrip"
    Q1 = Q1 & " ,Version"
    Q1 = Q1 & " ,OldCodigo"
    Q1 = Q1 & " ,(SELECT COUNT(*) FROM CodActiv) as Cant"
    Q1 = Q1 & " From CodActiv"
    Set Rs = OpenRsDao(DbAccess, Q1)

    If Rs.EOF = False Then
     If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CodActiv"
            Q1 = Q1 & " WHERE Codigo = " & vFldDao(Rs("Codigo"))
            Q1 = Q1 & " AND Descrip = " & vFldDao(Rs("Descrip"))
            Q1 = Q1 & " AND Version = " & vFldDao(Rs("Version"))
            Q1 = Q1 & " AND OldCodigo = " & vFldDao(Rs("OldCodigo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT CodActiv ON "
                Q1 = Q1 & " INSERT INTO CodActiv"
                Q1 = Q1 & " (Codigo"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Version"
                Q1 = Q1 & " ,OldCodigo)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " ('" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descrip")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Version"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("OldCodigo")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT CodActiv OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CodActiv"
                Q1 = Q1 & " SET Descrip = " & vFldDao(Rs("Descrip"))
                Q1 = Q1 & " ,Version = " & vFldDao(Rs("Version"))
                Q1 = Q1 & " ,OldCodigo = " & vFldDao(Rs("OldCodigo"))
                Q1 = Q1 & " WHERE Codigo = " & vFldDao(Rs("Codigo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasCartola(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Cartola' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Cartola ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdCartola"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,IdCuentaBco"
    Q1 = Q1 & " ,Cartola"
    Q1 = Q1 & " ,FDesde"
    Q1 = Q1 & " ,FHasta"
    Q1 = Q1 & " ,TotCargo"
    Q1 = Q1 & " ,TotAbono"
    Q1 = Q1 & " ,SaldoIni"
    Q1 = Q1 & " From Cartola"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Cartola"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCartola"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Cartola ON "
                Q1 = " INSERT INTO Cartola"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdCuentaBco"
                Q1 = Q1 & " ,Cartola"
                Q1 = Q1 & " ,FDesde"
                Q1 = Q1 & " ,FHasta"
                Q1 = Q1 & " ,TotCargo"
                Q1 = Q1 & " ,TotAbono"
                Q1 = Q1 & " ,SaldoIni"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaBco"))
                Q1 = Q1 & " ," & vFldDao(Rs("Cartola"))
                Q1 = Q1 & " ," & vFldDao(Rs("FDesde"))
                Q1 = Q1 & " ," & vFldDao(Rs("FHasta"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotCargo"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotAbono"))
                Q1 = Q1 & " ," & vFldDao(Rs("SaldoIni"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCartola")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Cartola OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Cartola"
                Q1 = Q1 & " SET IdCuentaBco = " & vFldDao(Rs("IdCuentaBco"))
                Q1 = Q1 & " ,Cartola = " & vFldDao(Rs("Cartola"))
                Q1 = Q1 & " ,FDesde = " & vFldDao(Rs("FDesde"))
                Q1 = Q1 & " ,FHasta = " & vFldDao(Rs("FHasta"))
                Q1 = Q1 & " ,TotCargo = " & vFldDao(Rs("TotCargo"))
                Q1 = Q1 & " ,TotAbono = " & vFldDao(Rs("TotAbono"))
                Q1 = Q1 & " ,SaldoIni = " & vFldDao(Rs("SaldoIni"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCartola"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
   
    Q1 = " UPDATE Cartola"
    Q1 = Q1 & " Set Cartola.IdCuentaBco = Cuentas.IdCuenta "
    Q1 = Q1 & " FROM Cartola INNER JOIN Cuentas ON Cartola.IdCuentaBco = Cuentas.IdTras AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCapPropioSimplAnual(DBSql As ADODB.Connection, DbAccess As Database)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim Rs2 As Recordset
    Dim CantSql As Long

    Q1 = "SELECT Count(IdCapPropioSimplAnual) as Cant"
    Q1 = Q1 & " From CapPropioSimplAnual"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

    Q1 = "SELECT IdCapPropioSimplAnual"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,TipoDetCPS"
    Q1 = Q1 & " ,IngresoManual"
    Q1 = Q1 & " ,AnoValor"
    Q1 = Q1 & " ,Valor"
    Q1 = Q1 & " ,(SELECT COUNT(*) FROM CapPropioSimplAnual) as Cant"
    Q1 = Q1 & " From CapPropioSimplAnual"
    Set Rs = OpenRsDao(DbAccess, Q1)
    
   If Rs.EOF = False Then
    
    If CantSql < vFldDao(Rs("Cant")) Then
    
        Do While Rs.EOF = False
             'Txt.Caption = "Tabla: CapPropioSimplAnual Registros: " & PG.Value & " De " & PG.Max
             Sleep (3000)
             
             Q1 = "SELECT * "
             Q1 = Q1 & " From CapPropioSimplAnual"
             Q1 = Q1 & " WHERE IdCapPropioSimplAnual = " & vFldDao(Rs("IdCapPropioSimplAnual"))
             Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
             Q1 = Q1 & " AND TipoDetCPS = " & vFldDao(Rs("TipoDetCPS"))
             Q1 = Q1 & " AND IngresoManual = " & vFldDao(Rs("IngresoManual"))
             Q1 = Q1 & " AND AnoValor = " & vFldDao(Rs("AnoValor"))
             Q1 = Q1 & " AND Valor = " & vFldDao(Rs("Valor"))
             Set Rs2 = OpenRs(DBSql, Q1)
             If Rs2.EOF = True Then
             
                 Q1 = " SET IDENTITY_INSERT CapPropioSimplAnual ON "
                 Q1 = Q1 & " INSERT INTO CapPropioSimplAnual"
                 Q1 = Q1 & " (IdCapPropioSimplAnual"
                 Q1 = Q1 & " ,IdEmpresa"
                 Q1 = Q1 & " ,TipoDetCPS"
                 Q1 = Q1 & " ,IngresoManual"
                 Q1 = Q1 & " ,AnoValor"
                 Q1 = Q1 & " ,Valor)"
                 Q1 = Q1 & " Values"
                 Q1 = Q1 & " (" & vFldDao(Rs("IdCapPropioSimplAnual"))
                 Q1 = Q1 & " ," & gEmpresa.id
                 Q1 = Q1 & " ," & vFldDao(Rs("TipoDetCPS"))
                 Q1 = Q1 & " ," & vFldDao(Rs("IngresoManual"))
                 Q1 = Q1 & " ," & vFldDao(Rs("AnoValor"))
                 Q1 = Q1 & " ," & vFldDao(Rs("Valor")) & ")"
                 Q1 = Q1 & " SET IDENTITY_INSERT CapPropioSimplAnual OFF  "
                 Call ExecSQL(DBSql, Q1)
                 
                 
             Else
             
                 Q1 = " UPDATE CapPropioSimplAnual"
                 Q1 = Q1 & " SET TipoDetCPS = " & vFldDao(Rs("TipoDetCPS"))
                 Q1 = Q1 & " ,IngresoManual = " & vFldDao(Rs("IngresoManual"))
                 Q1 = Q1 & " ,AnoValor = " & vFldDao(Rs("AnoValor"))
                 Q1 = Q1 & " ,Valor = " & vFldDao(Rs("Valor"))
                 Q1 = Q1 & " WHERE IdCapPropioSimplAnual = " & vFldDao(Rs("IdCapPropioSimplAnual"))
                 Q1 = Q1 & " AND   IdEmpresa = " & gEmpresa.id
                 Call ExecSQL(DBSql, Q1)
                 
                 
             End If
             Call CloseRs(Rs2)
     
           Rs.MoveNext
        Loop
        
     End If
    End If
   Call CloseRs(Rs)
   'Txt.Caption = "Tabla: CapPropioSimplAnual Traspasada Correctamente"


End Sub

Public Sub TrasBaseImponible14ter(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'BaseImponible14ter' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE BaseImponible14ter ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdBaseImponible14Ter"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,TipoBaseImp"
    Q1 = Q1 & " ,IdItemBaseImp"
    Q1 = Q1 & " ,Valor"
    Q1 = Q1 & " From BaseImponible14ter"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From BaseImponible14ter"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdBaseImponible14Ter"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT BaseImponible14ter ON "
                Q1 = " INSERT INTO BaseImponible14Ter"
                Q1 = Q1 & " (IdBaseImponible14Ter"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,TipoBaseImp"
                Q1 = Q1 & " ,IdItemBaseImp"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdBaseImponible14Ter"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoBaseImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdItemBaseImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Valor"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdBaseImponible14Ter")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT BaseImponible14ter OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE BaseImponible14ter"
                Q1 = Q1 & " SET TipoBaseImp = " & vFldDao(Rs("TipoBaseImp"))
                Q1 = Q1 & " ,IdItemBaseImp = " & vFldDao(Rs("IdItemBaseImp"))
                Q1 = Q1 & " ,Valor = " & vFldDao(Rs("Valor"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdBaseImponible14Ter"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasBaseImponible14D(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'BaseImponible14D' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE BaseImponible14D ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdBaseImponible14D"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,Tipo"
    Q1 = Q1 & " ,Nivel"
    Q1 = Q1 & " ,Codigo"
    Q1 = Q1 & " ,Fecha"
    Q1 = Q1 & " ,Valor"
    Q1 = Q1 & " From BaseImponible14D"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From BaseImponible14D"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdBaseImponible14D"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
              'Q1 = " SET IDENTITY_INSERT BaseImponible14D ON "
              Q1 = " INSERT INTO BaseImponible14D"
              Q1 = Q1 & " (IdEmpresa"
              Q1 = Q1 & " ,Ano"
              Q1 = Q1 & " ,Tipo"
              Q1 = Q1 & " ,Nivel"
              Q1 = Q1 & " ,Codigo"
              Q1 = Q1 & " ,Fecha"
              Q1 = Q1 & " ,Valor"
              Q1 = Q1 & " ,IdTras)"
              Q1 = Q1 & " Values"
              Q1 = Q1 & " (" & IdEmpresa
              Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
              Q1 = Q1 & " ," & vFldDao(Rs("Tipo"))
              Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
              Q1 = Q1 & " ," & vFldDao(Rs("Codigo"))
              Q1 = Q1 & " ," & vFldDao(Rs("Fecha"))
              Q1 = Q1 & " ," & vFldDao(Rs("Valor"))
              Q1 = Q1 & " ," & vFldDao(Rs("IdBaseImponible14D")) & ")"
              'Q1 = Q1 & " SET IDENTITY_INSERT BaseImponible14D OFF  "
              Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE BaseImponible14D"
                Q1 = Q1 & " SET Tipo = " & vFldDao(Rs("Tipo"))
                Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ,Codigo = " & vFldDao(Rs("Codigo"))
                Q1 = Q1 & " ,Fecha = " & vFldDao(Rs("Fecha"))
                Q1 = Q1 & " ,Valor = " & vFldDao(Rs("Valor"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdBaseImponible14D"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)

End Sub

Public Sub TrasAsistImpPrimCat(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'AsistImpPrimCat' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE AsistImpPrimCat ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdAsistImpPrimCat"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,IdItem"
    Q1 = Q1 & " ,RemEjAntNominal"
    Q1 = Q1 & " ,RemEjAntAct"
    Q1 = Q1 & " ,GeneradoAno"
    Q1 = Q1 & " ,CredUtilizado"
    Q1 = Q1 & " ,RemEjSgte"
    Q1 = Q1 & " From AsistImpPrimCat"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AsistImpPrimCat"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAsistImpPrimCat"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT AsistImpPrimCat ON "
                Q1 = " INSERT INTO AsistImpPrimCat"
                Q1 = Q1 & " (IdAsistImpPrimCat"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdItem"
                Q1 = Q1 & " ,RemEjAntNominal"
                Q1 = Q1 & " ,RemEjAntAct"
                Q1 = Q1 & " ,GeneradoAno"
                Q1 = Q1 & " ,CredUtilizado"
                Q1 = Q1 & " ,RemEjSgte"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdAsistImpPrimCat"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdItem"))
                Q1 = Q1 & " ," & vFldDao(Rs("RemEjAntNominal"))
                Q1 = Q1 & " ," & vFldDao(Rs("RemEjAntAct"))
                Q1 = Q1 & " ," & vFldDao(Rs("GeneradoAno"))
                Q1 = Q1 & " ," & vFldDao(Rs("CredUtilizado"))
                Q1 = Q1 & " ," & vFldDao(Rs("RemEjSgte"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAsistImpPrimCat")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AsistImpPrimCat OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AsistImpPrimCat"
                Q1 = Q1 & " SET IdItem = " & vFldDao(Rs("IdItem"))
                Q1 = Q1 & " ,RemEjAntNominal = " & vFldDao(Rs("RemEjAntNominal"))
                Q1 = Q1 & " ,RemEjAntAct = " & vFldDao(Rs("RemEjAntAct"))
                Q1 = Q1 & " ,GeneradoAno = " & vFldDao(Rs("GeneradoAno"))
                Q1 = Q1 & " ,CredUtilizado = " & vFldDao(Rs("CredUtilizado"))
                Q1 = Q1 & " ,RemEjSgte = " & vFldDao(Rs("RemEjSgte"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAsistImpPrimCat"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasAreaNegocio(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'AreaNegocio' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE AreaNegocio ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdAreaNegocio"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Codigo"
    Q1 = Q1 & " ,Descripcion"
    Q1 = Q1 & " ,Vigente"
    Q1 = Q1 & " From AreaNegocio"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AreaNegocio"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAreaNegocio"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT AreaNegocio ON "
                Q1 = " INSERT INTO AreaNegocio"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,Vigente"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAreaNegocio")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AreaNegocio OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AreaNegocio"
                Q1 = Q1 & " SET Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,Vigente = " & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAreaNegocio"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasPlanBasico(DBSql As ADODB.Connection, DbAccess As Database)

   Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM PlanBasico"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From PlanBasico"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idCuenta"
   Q1 = Q1 & " ,idPadre"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,CodFECU"
   Q1 = Q1 & " ,Nivel"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,MarcaApertura"
   Q1 = Q1 & " ,TipoCapPropio"
   Q1 = Q1 & " ,CodF22"
   Q1 = Q1 & " ,Atrib1"
   Q1 = Q1 & " ,Atrib2"
   Q1 = Q1 & " ,Atrib3"
   Q1 = Q1 & " ,Atrib4"
   Q1 = Q1 & " ,Atrib5"
   Q1 = Q1 & " ,Atrib6"
   Q1 = Q1 & " ,Atrib7"
   Q1 = Q1 & " ,Atrib8"
   Q1 = Q1 & " ,Atrib9"
   Q1 = Q1 & " ,Atrib10"
   Q1 = Q1 & " ,CodIFRS_EstRes"
   Q1 = Q1 & " ,CodIFRS_EstFin"
   Q1 = Q1 & " ,CodIFRS"
   Q1 = Q1 & " ,TipoPartida"
   Q1 = Q1 & " ,CodCtaPlanSII"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM PlanBasico) as Cant"
   Q1 = Q1 & " From PlanBasico"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
       If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From PlanBasico"
            Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT PlanBasico ON "
                Q1 = Q1 & " INSERT INTO PlanBasico"
                Q1 = Q1 & " (idCuenta"
                Q1 = Q1 & " ,idPadre"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,CodFECU"
                Q1 = Q1 & " ,Nivel"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Clasificacion"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,MarcaApertura"
                Q1 = Q1 & " ,TipoCapPropio"
                Q1 = Q1 & " ,CodF22"
                Q1 = Q1 & " ,Atrib1"
                Q1 = Q1 & " ,Atrib2"
                Q1 = Q1 & " ,Atrib3"
                Q1 = Q1 & " ,Atrib4"
                Q1 = Q1 & " ,Atrib5"
                Q1 = Q1 & " ,Atrib6"
                Q1 = Q1 & " ,Atrib7"
                Q1 = Q1 & " ,Atrib8"
                Q1 = Q1 & " ,Atrib9"
                Q1 = Q1 & " ,Atrib10"
                Q1 = Q1 & " ,CodIFRS_EstRes"
                Q1 = Q1 & " ,CodIFRS_EstFin"
                Q1 = Q1 & " ,CodIFRS"
                Q1 = Q1 & " ,TipoPartida"
                Q1 = Q1 & " ,CodCtaPlanSII)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ," & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ," & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaPlanSII")) & "')"
                Q1 = Q1 & " SET IDENTITY_INSERT PlanBasico OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE PlanBasico"
                Q1 = Q1 & " SET idPadre = " & vFldDao(Rs("idPadre"))
                Q1 = Q1 & " ,Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,CodFECU = '" & vFldDao(Rs("CodFECU")) & "'"
                Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,MarcaApertura = " & vFldDao(Rs("MarcaApertura"))
                Q1 = Q1 & " ,TipoCapPropio = " & vFldDao(Rs("TipoCapPropio"))
                Q1 = Q1 & " ,CodF22 = " & vFldDao(Rs("CodF22"))
                Q1 = Q1 & " ,Atrib1 = " & vFldDao(Rs("Atrib1"))
                Q1 = Q1 & " ,Atrib2 = " & vFldDao(Rs("Atrib2"))
                Q1 = Q1 & " ,Atrib3 = " & vFldDao(Rs("Atrib3"))
                Q1 = Q1 & " ,Atrib4 = " & vFldDao(Rs("Atrib4"))
                Q1 = Q1 & " ,Atrib5 = " & vFldDao(Rs("Atrib5"))
                Q1 = Q1 & " ,Atrib6 = " & vFldDao(Rs("Atrib6"))
                Q1 = Q1 & " ,Atrib7 = " & vFldDao(Rs("Atrib7"))
                Q1 = Q1 & " ,Atrib8 = " & vFldDao(Rs("Atrib8"))
                Q1 = Q1 & " ,Atrib9 = " & vFldDao(Rs("Atrib9"))
                Q1 = Q1 & " ,Atrib10 = " & vFldDao(Rs("Atrib10"))
                Q1 = Q1 & " ,CodIFRS_EstRes = '" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                Q1 = Q1 & " ,CodIFRS_EstFin = '" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                Q1 = Q1 & " ,CodIFRS = '" & vFldDao(Rs("CodIFRS")) & "'"
                Q1 = Q1 & " ,TipoPartida = " & vFldDao(Rs("TipoPartida"))
                Q1 = Q1 & " ,CodCtaPlanSII = '" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
   End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasPlanAvanzado(DBSql As ADODB.Connection, DbAccess As Database)

   Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "DELETE FROM PlanAvanzado"
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT Count(*) as Cant"
    Q1 = Q1 & " From PlanAvanzado"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idCuenta"
   Q1 = Q1 & " ,idPadre"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,CodFECU"
   Q1 = Q1 & " ,Nivel"
   Q1 = Q1 & " ,Estado"
   Q1 = Q1 & " ,Clasificacion"
   Q1 = Q1 & " ,Debe"
   Q1 = Q1 & " ,Haber"
   Q1 = Q1 & " ,MarcaApertura"
   Q1 = Q1 & " ,TipoCapPropio"
   Q1 = Q1 & " ,CodF22"
   Q1 = Q1 & " ,Atrib1"
   Q1 = Q1 & " ,Atrib2"
   Q1 = Q1 & " ,Atrib3"
   Q1 = Q1 & " ,Atrib4"
   Q1 = Q1 & " ,Atrib5"
   Q1 = Q1 & " ,Atrib6"
   Q1 = Q1 & " ,Atrib7"
   Q1 = Q1 & " ,Atrib8"
   Q1 = Q1 & " ,Atrib9"
   Q1 = Q1 & " ,Atrib10"
   Q1 = Q1 & " ,CodIFRS_EstRes"
   Q1 = Q1 & " ,CodIFRS_EstFin"
   Q1 = Q1 & " ,CodIFRS"
   Q1 = Q1 & " ,TipoPartida"
   Q1 = Q1 & " ,CodCtaPlanSII"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM PlanAvanzado) as Cant"
   Q1 = Q1 & " From PlanAvanzado"
   Set Rs = OpenRsDao(DbAccess, Q1)
        
   If Rs.EOF = False Then
   
           If CantSql < vFldDao(Rs("Cant")) Then
       
           Do While Rs.EOF = False
           
                Q1 = "SELECT * "
                Q1 = Q1 & " From PlanAvanzado"
                Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
                Set Rs1 = OpenRs(DBSql, Q1)
                
                If Rs1.EOF = True Then
                
                    Q1 = " SET IDENTITY_INSERT PlanAvanzado ON "
                    Q1 = Q1 & " INSERT INTO PlanAvanzado"
                    Q1 = Q1 & " (idCuenta"
                    Q1 = Q1 & " ,idPadre"
                    Q1 = Q1 & " ,Codigo"
                    Q1 = Q1 & " ,Nombre"
                    Q1 = Q1 & " ,Descripcion"
                    Q1 = Q1 & " ,CodFECU"
                    Q1 = Q1 & " ,Nivel"
                    Q1 = Q1 & " ,Estado"
                    Q1 = Q1 & " ,Clasificacion"
                    Q1 = Q1 & " ,Debe"
                    Q1 = Q1 & " ,Haber"
                    Q1 = Q1 & " ,MarcaApertura"
                    Q1 = Q1 & " ,TipoCapPropio"
                    Q1 = Q1 & " ,CodF22"
                    Q1 = Q1 & " ,Atrib1"
                    Q1 = Q1 & " ,Atrib2"
                    Q1 = Q1 & " ,Atrib3"
                    Q1 = Q1 & " ,Atrib4"
                    Q1 = Q1 & " ,Atrib5"
                    Q1 = Q1 & " ,Atrib6"
                    Q1 = Q1 & " ,Atrib7"
                    Q1 = Q1 & " ,Atrib8"
                    Q1 = Q1 & " ,Atrib9"
                    Q1 = Q1 & " ,Atrib10"
                    Q1 = Q1 & " ,CodIFRS_EstRes"
                    Q1 = Q1 & " ,CodIFRS_EstFin"
                    Q1 = Q1 & " ,CodIFRS"
                    Q1 = Q1 & " ,TipoPartida"
                    Q1 = Q1 & " ,CodCtaPlanSII)"
                    Q1 = Q1 & " Values"
                    Q1 = Q1 & " (" & vFldDao(Rs("idCuenta"))
                    Q1 = Q1 & " ," & vFldDao(Rs("idPadre"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodFECU")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("Nivel"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Clasificacion"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                    Q1 = Q1 & " ," & vFldDao(Rs("MarcaApertura"))
                    Q1 = Q1 & " ," & vFldDao(Rs("TipoCapPropio"))
                    Q1 = Q1 & " ," & vFldDao(Rs("CodF22"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib1"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib2"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib3"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib4"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib5"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib6"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib7"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib8"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib9"))
                    Q1 = Q1 & " ," & vFldDao(Rs("Atrib10"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodIFRS")) & "'"
                    Q1 = Q1 & " ," & vFldDao(Rs("TipoPartida"))
                    Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaPlanSII")) & "')"
                    Q1 = Q1 & " SET IDENTITY_INSERT PlanAvanzado OFF  "
                    Call ExecSQL(DBSql, Q1)
                    
                Else
                
                    Q1 = " UPDATE PlanAvanzado"
                    Q1 = Q1 & " SET idPadre = " & vFldDao(Rs("idPadre"))
                    Q1 = Q1 & " ,Codigo = '" & vFldDao(Rs("Codigo")) & "'"
                    Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                    Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                    Q1 = Q1 & " ,CodFECU = '" & vFldDao(Rs("CodFECU")) & "'"
                    Q1 = Q1 & " ,Nivel = " & vFldDao(Rs("Nivel"))
                    Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                    Q1 = Q1 & " ,Clasificacion = " & vFldDao(Rs("Clasificacion"))
                    Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                    Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                    Q1 = Q1 & " ,MarcaApertura = " & vFldDao(Rs("MarcaApertura"))
                    Q1 = Q1 & " ,TipoCapPropio = " & vFldDao(Rs("TipoCapPropio"))
                    Q1 = Q1 & " ,CodF22 = " & vFldDao(Rs("CodF22"))
                    Q1 = Q1 & " ,Atrib1 = " & vFldDao(Rs("Atrib1"))
                    Q1 = Q1 & " ,Atrib2 = " & vFldDao(Rs("Atrib2"))
                    Q1 = Q1 & " ,Atrib3 = " & vFldDao(Rs("Atrib3"))
                    Q1 = Q1 & " ,Atrib4 = " & vFldDao(Rs("Atrib4"))
                    Q1 = Q1 & " ,Atrib5 = " & vFldDao(Rs("Atrib5"))
                    Q1 = Q1 & " ,Atrib6 = " & vFldDao(Rs("Atrib6"))
                    Q1 = Q1 & " ,Atrib7 = " & vFldDao(Rs("Atrib7"))
                    Q1 = Q1 & " ,Atrib8 = " & vFldDao(Rs("Atrib8"))
                    Q1 = Q1 & " ,Atrib9 = " & vFldDao(Rs("Atrib9"))
                    Q1 = Q1 & " ,Atrib10 = " & vFldDao(Rs("Atrib10"))
                    Q1 = Q1 & " ,CodIFRS_EstRes = '" & vFldDao(Rs("CodIFRS_EstRes")) & "'"
                    Q1 = Q1 & " ,CodIFRS_EstFin = '" & vFldDao(Rs("CodIFRS_EstFin")) & "'"
                    Q1 = Q1 & " ,CodIFRS = '" & vFldDao(Rs("CodIFRS")) & "'"
                    Q1 = Q1 & " ,TipoPartida = " & vFldDao(Rs("TipoPartida"))
                    Q1 = Q1 & " ,CodCtaPlanSII = '" & vFldDao(Rs("CodCtaPlanSII")) & "'"
                    Q1 = Q1 & " WHERE idCuenta = " & vFldDao(Rs("idCuenta"))
                    Call ExecSQL(DBSql, Q1)
                    
                End If
                Call CloseRs(Rs1)
    
          Rs.MoveNext
          Loop
       End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasSocios(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Socios' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Socios ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdSocio"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,RUT"
   Q1 = Q1 & " ,Nombre"
   Q1 = Q1 & " ,PjePart"
   Q1 = Q1 & " ,MontoSuscrito"
   Q1 = Q1 & " ,MontoPagado"
   Q1 = Q1 & " ,IdCuentaAportes"
   Q1 = Q1 & " ,IdCuentaRetiros"
   Q1 = Q1 & " ,IdTipoSocio"
   Q1 = Q1 & " ,Vigente"
   Q1 = Q1 & " ,CantAcciones"
   Q1 = Q1 & " ,MontoIngresadoUsuario"
   Q1 = Q1 & " ,MontoATraspasar"
   Q1 = Q1 & " From Socios"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Socios"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdSocio"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT Socios ON "
                Q1 = " INSERT INTO Socios"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,RUT"
                Q1 = Q1 & " ,Nombre"
                Q1 = Q1 & " ,PjePart"
                Q1 = Q1 & " ,MontoSuscrito"
                Q1 = Q1 & " ,MontoPagado"
                Q1 = Q1 & " ,IdCuentaAportes"
                Q1 = Q1 & " ,IdCuentaRetiros"
                Q1 = Q1 & " ,IdTipoSocio"
                Q1 = Q1 & " ,Vigente"
                Q1 = Q1 & " ,CantAcciones"
                Q1 = Q1 & " ,MontoIngresadoUsuario"
                Q1 = Q1 & " ,MontoATraspasar"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("RUT")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("PjePart"))))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoSuscrito"))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoPagado"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaAportes"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaRetiros"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdTipoSocio"))
                Q1 = Q1 & " ," & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " ," & vFldDao(Rs("CantAcciones"))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoIngresadoUsuario"))
                Q1 = Q1 & " ," & vFldDao(Rs("MontoATraspasar"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdSocio")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT Socios OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Socios"
                Q1 = Q1 & " SET RUT = '" & vFldDao(Rs("RUT")) & "'"
                Q1 = Q1 & " ,Nombre = '" & vFldDao(Rs("Nombre")) & "'"
                Q1 = Q1 & " ,PjePart = " & str(vFmt(vFldDao(Rs("PjePart"))))
                Q1 = Q1 & " ,MontoSuscrito = " & vFldDao(Rs("MontoSuscrito"))
                Q1 = Q1 & " ,MontoPagado = " & vFldDao(Rs("MontoPagado"))
                Q1 = Q1 & " ,IdCuentaAportes = " & vFldDao(Rs("IdCuentaAportes"))
                Q1 = Q1 & " ,IdCuentaRetiros = " & vFldDao(Rs("IdCuentaRetiros"))
                Q1 = Q1 & " ,IdTipoSocio = " & vFldDao(Rs("IdTipoSocio"))
                Q1 = Q1 & " ,Vigente = " & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " ,CantAcciones = " & vFldDao(Rs("CantAcciones"))
                Q1 = Q1 & " ,MontoIngresadoUsuario = " & vFldDao(Rs("MontoIngresadoUsuario"))
                Q1 = Q1 & " ,MontoATraspasar = " & vFldDao(Rs("MontoATraspasar"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdSocio"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    
    Q1 = "UPDATE S"
    Q1 = Q1 & " SET S.IdCuentaAportes = ISNULL(C.idCuenta,S.IdCuentaAportes)"
    Q1 = Q1 & " FROM Socios S"
    Q1 = Q1 & " INNER JOIN Cuentas C ON C.IdTras = S.IdCuentaAportes AND C.IdEmpresa = S.IdEmpresa AND C.Ano = S.Ano"
    Q1 = Q1 & " Where s.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND S.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "UPDATE S"
    Q1 = Q1 & " SET S.IdCuentaRetiros = ISNULL(CU.IdCuenta, S.IdCuentaRetiros)"
    Q1 = Q1 & " FROM Socios S"
    Q1 = Q1 & " INNER JOIN Cuentas CU ON CU.IdTras = S.IdCuentaRetiros AND CU.IdEmpresa = S.IdEmpresa AND CU.Ano = S.Ano"
    Q1 = Q1 & " Where s.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND S.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasCentroCosto(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
   Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'CentroCosto' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE CentroCosto ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdCCosto"
   Q1 = Q1 & " ,IdEmpresa"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Descripcion"
   Q1 = Q1 & " ,Vigente"
   Q1 = Q1 & " From CentroCosto"
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
    
        Q1 = Q1 & " DELETE FROM CentroCosto  "
        Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
        Call ExecSQL(DBSql, Q1)
    
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From CentroCosto"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCCosto"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CentroCosto ON "
                Q1 = " INSERT INTO CentroCosto"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Descripcion"
                Q1 = Q1 & " ,Vigente"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("Codigo")) & "'" 'str(vFmt(vFldDao(Rs("Codigo")))) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCCosto")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CentroCosto OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE CentroCosto"
                Q1 = Q1 & " SET Codigo = '" & str(vFmt(vFldDao(Rs("Codigo")))) & "'"
                Q1 = Q1 & " ,Descripcion = '" & vFldDao(Rs("Descripcion")) & "'"
                Q1 = Q1 & " ,Vigente = " & vFldDao(Rs("Vigente"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCCosto"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
      End If

   Call CloseRs(Rs)


End Sub

Public Sub TrasTipoValor(DBSql As ADODB.Connection, DbAccess As Database)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = Q1 & " DELETE FROM TipoValor  "
    Call ExecSQL(DBSql, Q1)
   

   Q1 = "SELECT idTValor"
   Q1 = Q1 & " ,TipoLib"
   Q1 = Q1 & " ,Codigo"
   Q1 = Q1 & " ,Valor"
   Q1 = Q1 & " ,Diminutivo"
   Q1 = Q1 & " ,Atributo"
   Q1 = Q1 & " ,Multiple"
   Q1 = Q1 & " ,CodF29"
   Q1 = Q1 & " ,CodF29_Adic"
   Q1 = Q1 & " ,TipoDoc"
   Q1 = Q1 & " ,Tit1"
   Q1 = Q1 & " ,Tit2"
   Q1 = Q1 & " ,CodImpSII"
   Q1 = Q1 & " ,Orden"
   Q1 = Q1 & " ,Tasa"
   Q1 = Q1 & " ,EsRecuperable"
   Q1 = Q1 & " ,CodSIIDTE"
   Q1 = Q1 & " ,TitCompleto"
   Q1 = Q1 & " ,TipoIVARetenido"
   Q1 = Q1 & " From TipoValor"
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From TipoValor"
            Q1 = Q1 & " WHERE idTValor = " & vFldDao(Rs("idTValor"))
            Q1 = Q1 & " AND TipoLib = " & vFldDao(Rs("TipoLib"))
            Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " SET IDENTITY_INSERT TipoValor ON "
                Q1 = Q1 & " INSERT INTO TipoValor"
                Q1 = Q1 & " (idTValor"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,Codigo"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,Diminutivo"
                Q1 = Q1 & " ,Atributo"
                Q1 = Q1 & " ,Multiple"
                Q1 = Q1 & " ,CodF29"
                Q1 = Q1 & " ,CodF29_Adic"
                Q1 = Q1 & " ,TipoDoc"
                Q1 = Q1 & " ,Tit1"
                Q1 = Q1 & " ,Tit2"
                Q1 = Q1 & " ,CodImpSII"
                Q1 = Q1 & " ,Orden"
                Q1 = Q1 & " ,Tasa"
                Q1 = Q1 & " ,EsRecuperable"
                Q1 = Q1 & " ,CodSIIDTE"
                Q1 = Q1 & " ,TitCompleto"
                Q1 = Q1 & " ,TipoIVARetenido)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("idTValor"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("Codigo"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("Valor")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Atributo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Multiple"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodF29_Adic"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("TipoDoc")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Tit1")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("Tit2")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodImpSII")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Tasa"))))
                Q1 = Q1 & " ," & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodSIIDTE")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("TitCompleto")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("TipoIVARetenido")) & ")"
                Q1 = Q1 & " SET IDENTITY_INSERT TipoValor OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE TipoValor"
                Q1 = Q1 & " SET Valor = '" & vFldDao(Rs("Valor")) & "'"
                Q1 = Q1 & " ,Diminutivo = '" & vFldDao(Rs("Diminutivo")) & "'"
                Q1 = Q1 & " ,Atributo = '" & vFldDao(Rs("Atributo")) & "'"
                Q1 = Q1 & " ,Multiple = " & vFldDao(Rs("Multiple"))
                Q1 = Q1 & " ,CodF29 = " & vFldDao(Rs("CodF29"))
                Q1 = Q1 & " ,CodF29_Adic = " & vFldDao(Rs("CodF29_Adic"))
                Q1 = Q1 & " ,TipoDoc = '" & vFldDao(Rs("TipoDoc")) & "'"
                Q1 = Q1 & " ,Tit1 = '" & vFldDao(Rs("Tit1")) & "'"
                Q1 = Q1 & " ,Tit2 = '" & vFldDao(Rs("Tit2")) & "'"
                Q1 = Q1 & " ,CodImpSII = '" & vFldDao(Rs("CodImpSII")) & "'"
                Q1 = Q1 & " ,Orden = " & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ,Tasa = " & str(vFmt(vFldDao(Rs("Tasa"))))
                Q1 = Q1 & " ,EsRecuperable = " & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,CodSIIDTE = '" & vFldDao(Rs("CodSIIDTE")) & "'"
                Q1 = Q1 & " ,TitCompleto = '" & vFldDao(Rs("TitCompleto")) & "'"
                Q1 = Q1 & " ,TipoIVARetenido = " & vFldDao(Rs("TipoIVARetenido"))
                Q1 = Q1 & " WHERE idTValor = " & vFldDao(Rs("idTValor"))
                Q1 = Q1 & " AND TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " AND Codigo = " & vFldDao(Rs("Codigo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)


End Sub

Public Sub TrasEmpresasAno(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

    Dim Rs As dao.Recordset
    Dim Rs1 As Recordset
    Dim CantSql As Long
    
    Q1 = "SELECT Count(IdEmpresa) as Cant"
    Q1 = Q1 & " From EmpresasAno"
    Set Rs1 = OpenRs(DBSql, Q1)
    
    If Rs1.EOF = False Then
        CantSql = vFld(Rs1("Cant"))
    End If
    Call CloseRs(Rs1)

   Q1 = "SELECT idEmpresa"
   Q1 = Q1 & " ,Ano"
   Q1 = Q1 & " ,FCierre"
   Q1 = Q1 & " ,FApertura"
   Q1 = Q1 & " ,NCompAper"
   Q1 = Q1 & " ,NCompAperProx"
   Q1 = Q1 & " ,IdCompAper"
   Q1 = Q1 & " ,NumLastCompUnico"
   Q1 = Q1 & " ,NumLastCompA"
   Q1 = Q1 & " ,NumLastCompE"
   Q1 = Q1 & " ,NumLastCompI"
   Q1 = Q1 & " ,NumLastCompT"
   Q1 = Q1 & " ,NCompAperTrib"
   Q1 = Q1 & " ,IdCompAperTrib"
   Q1 = Q1 & " ,RemIVAUTM"
   Q1 = Q1 & " ,RemIVAUTMAnoAnt"
   Q1 = Q1 & " ,SaldoLibroCaja"
   Q1 = Q1 & " ,CredArt33bis"
   Q1 = Q1 & " ,CPS_CapitalAportado"
   Q1 = Q1 & " ,CPS_BaseImpPrimCat_14DN3"
   Q1 = Q1 & " ,CPS_BaseImpPrimCat_14DN8"
   Q1 = Q1 & " ,CPS_Participaciones"
   Q1 = Q1 & " ,CPS_Disminuciones"
   Q1 = Q1 & " ,CPS_GastosRechazados"
   Q1 = Q1 & " ,CPS_RetirosDividendos"
   Q1 = Q1 & " ,CPS_CapPropioSimplificado"
   Q1 = Q1 & " ,CPS_AumentosCapital"
   Q1 = Q1 & " ,CPS_GastosRechazadosNoPagan40"
   Q1 = Q1 & " ,CPS_INRPropios"
   Q1 = Q1 & " ,CPS_OtrosAjustesAumentos"
   Q1 = Q1 & " ,CPS_OtrosAjustesDisminuciones"
   Q1 = Q1 & " ,CPS_CapPropioTrib"
   Q1 = Q1 & " ,CPS_CapPropioTribAnoAnt"
   Q1 = Q1 & " ,CPS_RepPerdidaArrastre"
   Q1 = Q1 & " ,CPS_CapPropioSimplVarAnual"
   Q1 = Q1 & " ,CPS_INRPropiosPerdidas"
   Q1 = Q1 & " ,CPS_UtilidadesPerdida"
   Q1 = Q1 & " ,CPS_IngresoDiferido"
   Q1 = Q1 & " ,CPS_CTDImputableIPE"
   Q1 = Q1 & " ,CPS_IncentivoAhorro"
   Q1 = Q1 & " ,CPS_IDPCVoluntario"
   Q1 = Q1 & " ,CPS_CredActFijos"
   Q1 = Q1 & " ,CPS_CredParticipaciones"
   Q1 = Q1 & " ,(SELECT COUNT(*) FROM EmpresasAno) as Cant"
   Q1 = Q1 & " From EmpresasAno"
   Q1 = Q1 & " WHERE idEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   If Rs.EOF = False Then
    If CantSql < vFldDao(Rs("Cant")) Then
   
       Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From EmpresasAno"
            Q1 = Q1 & " WHERE idEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
                'Q1 = " SET IDENTITY_INSERT EmpresasAno ON "
                Q1 = " INSERT INTO EmpresasAno"
                Q1 = Q1 & " (idEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,FCierre"
                Q1 = Q1 & " ,FApertura"
                Q1 = Q1 & " ,NCompAper"
                Q1 = Q1 & " ,NCompAperProx"
                Q1 = Q1 & " ,IdCompAper"
                Q1 = Q1 & " ,NumLastCompUnico"
                Q1 = Q1 & " ,NumLastCompA"
                Q1 = Q1 & " ,NumLastCompE"
                Q1 = Q1 & " ,NumLastCompI"
                Q1 = Q1 & " ,NumLastCompT"
                Q1 = Q1 & " ,NCompAperTrib"
                Q1 = Q1 & " ,IdCompAperTrib"
                Q1 = Q1 & " ,RemIVAUTM"
                Q1 = Q1 & " ,RemIVAUTMAnoAnt"
                Q1 = Q1 & " ,SaldoLibroCaja"
                Q1 = Q1 & " ,CredArt33bis"
                Q1 = Q1 & " ,CPS_CapitalAportado"
                Q1 = Q1 & " ,CPS_BaseImpPrimCat_14DN3"
                Q1 = Q1 & " ,CPS_BaseImpPrimCat_14DN8"
                Q1 = Q1 & " ,CPS_Participaciones"
                Q1 = Q1 & " ,CPS_Disminuciones"
                Q1 = Q1 & " ,CPS_GastosRechazados"
                Q1 = Q1 & " ,CPS_RetirosDividendos"
                Q1 = Q1 & " ,CPS_CapPropioSimplificado"
                Q1 = Q1 & " ,CPS_AumentosCapital"
                Q1 = Q1 & " ,CPS_GastosRechazadosNoPagan40"
                Q1 = Q1 & " ,CPS_INRPropios"
                Q1 = Q1 & " ,CPS_OtrosAjustesAumentos"
                Q1 = Q1 & " ,CPS_OtrosAjustesDisminuciones"
                Q1 = Q1 & " ,CPS_CapPropioTrib"
                Q1 = Q1 & " ,CPS_CapPropioTribAnoAnt"
                Q1 = Q1 & " ,CPS_RepPerdidaArrastre"
                Q1 = Q1 & " ,CPS_CapPropioSimplVarAnual"
                Q1 = Q1 & " ,CPS_INRPropiosPerdidas"
                Q1 = Q1 & " ,CPS_UtilidadesPerdida"
                Q1 = Q1 & " ,CPS_IngresoDiferido"
                Q1 = Q1 & " ,CPS_CTDImputableIPE"
                Q1 = Q1 & " ,CPS_IncentivoAhorro"
                Q1 = Q1 & " ,CPS_IDPCVoluntario"
                Q1 = Q1 & " ,CPS_CredActFijos"
                Q1 = Q1 & " ,CPS_CredParticipaciones)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("FCierre"))
                Q1 = Q1 & " ," & vFldDao(Rs("FApertura"))
                Q1 = Q1 & " ," & vFldDao(Rs("NCompAper"))
                Q1 = Q1 & " ," & vFldDao(Rs("NCompAperProx"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompAper"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumLastCompUnico"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumLastCompA"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumLastCompE"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumLastCompI"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumLastCompT"))
                Q1 = Q1 & " ," & vFldDao(Rs("NCompAperTrib"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompAperTrib"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("RemIVAUTM"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("RemIVAUTMAnoAnt"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("SaldoLibroCaja"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CredArt33bis"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CapitalAportado"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_BaseImpPrimCat_14DN3"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_BaseImpPrimCat_14DN8"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_Participaciones"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_Disminuciones"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_GastosRechazados"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_RetirosDividendos"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CapPropioSimplificado"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_AumentosCapital"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_GastosRechazadosNoPagan40"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_INRPropios"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_OtrosAjustesAumentos"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_OtrosAjustesDisminuciones"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CapPropioTrib"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CapPropioTribAnoAnt"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_RepPerdidaArrastre"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CapPropioSimplVarAnual"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_INRPropiosPerdidas"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_UtilidadesPerdida"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_IngresoDiferido"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CTDImputableIPE"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_IncentivoAhorro"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_IDPCVoluntario"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CredActFijos"))))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("CPS_CredParticipaciones")))) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT EmpresasAno OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE EmpresasAno"
                Q1 = Q1 & " SET FCierre = " & vFldDao(Rs("FCierre"))
                Q1 = Q1 & " , FApertura = " & vFldDao(Rs("FApertura"))
                Q1 = Q1 & " , NCompAper = " & vFldDao(Rs("NCompAper"))
                Q1 = Q1 & " , NCompAperProx = " & vFldDao(Rs("NCompAperProx"))
                Q1 = Q1 & " , IdCompAper = " & vFldDao(Rs("IdCompAper"))
                Q1 = Q1 & " , NumLastCompUnico = " & vFldDao(Rs("NumLastCompUnico"))
                Q1 = Q1 & " , NumLastCompA = " & vFldDao(Rs("NumLastCompA"))
                Q1 = Q1 & " , NumLastCompE = " & vFldDao(Rs("NumLastCompE"))
                Q1 = Q1 & " , NumLastCompI = " & vFldDao(Rs("NumLastCompI"))
                Q1 = Q1 & " , NumLastCompT = " & vFldDao(Rs("NumLastCompT"))
                Q1 = Q1 & " , NCompAperTrib = " & vFldDao(Rs("NCompAperTrib"))
                Q1 = Q1 & " , IdCompAperTrib = " & vFldDao(Rs("IdCompAperTrib"))
                Q1 = Q1 & " , RemIVAUTM = " & vFldDao(Rs("RemIVAUTM"))
                Q1 = Q1 & " , RemIVAUTMAnoAnt = " & vFldDao(Rs("RemIVAUTMAnoAnt"))
                Q1 = Q1 & " , SaldoLibroCaja = " & vFldDao(Rs("SaldoLibroCaja"))
                Q1 = Q1 & " , CredArt33bis = " & vFldDao(Rs("CredArt33bis"))
                Q1 = Q1 & " , CPS_CapitalAportado = " & vFldDao(Rs("CPS_CapitalAportado"))
                Q1 = Q1 & " , CPS_BaseImpPrimCat_14DN3 = " & vFldDao(Rs("CPS_BaseImpPrimCat_14DN3"))
                Q1 = Q1 & " , CPS_BaseImpPrimCat_14DN8 = " & vFldDao(Rs("CPS_BaseImpPrimCat_14DN8"))
                Q1 = Q1 & " , CPS_Participaciones = " & vFldDao(Rs("CPS_Participaciones"))
                Q1 = Q1 & " , CPS_Disminuciones = " & vFldDao(Rs("CPS_Disminuciones"))
                Q1 = Q1 & " , CPS_GastosRechazados = " & vFldDao(Rs("CPS_GastosRechazados"))
                Q1 = Q1 & " , CPS_RetirosDividendos = " & vFldDao(Rs("CPS_RetirosDividendos"))
                Q1 = Q1 & " , CPS_CapPropioSimplificado = " & vFldDao(Rs("CPS_CapPropioSimplificado"))
                Q1 = Q1 & " , CPS_AumentosCapital = " & vFldDao(Rs("CPS_AumentosCapital"))
                Q1 = Q1 & " , CPS_GastosRechazadosNoPagan40 = " & vFldDao(Rs("CPS_GastosRechazadosNoPagan40"))
                Q1 = Q1 & " , CPS_INRPropios = " & vFldDao(Rs("CPS_INRPropios"))
                Q1 = Q1 & " , CPS_OtrosAjustesAumentos = " & vFldDao(Rs("CPS_OtrosAjustesAumentos"))
                Q1 = Q1 & " , CPS_OtrosAjustesDisminuciones = " & vFldDao(Rs("CPS_OtrosAjustesDisminuciones"))
                Q1 = Q1 & " , CPS_CapPropioTrib = " & vFldDao(Rs("CPS_CapPropioTrib"))
                Q1 = Q1 & " , CPS_CapPropioTribAnoAnt = " & vFldDao(Rs("CPS_CapPropioTribAnoAnt"))
                Q1 = Q1 & " , CPS_RepPerdidaArrastre = " & vFldDao(Rs("CPS_RepPerdidaArrastre"))
                Q1 = Q1 & " , CPS_CapPropioSimplVarAnual = " & vFldDao(Rs("CPS_CapPropioSimplVarAnual"))
                Q1 = Q1 & " , CPS_INRPropiosPerdidas = " & vFldDao(Rs("CPS_INRPropiosPerdidas"))
                Q1 = Q1 & " , CPS_UtilidadesPerdida = " & vFldDao(Rs("CPS_UtilidadesPerdida"))
                Q1 = Q1 & " , CPS_IngresoDiferido = " & vFldDao(Rs("CPS_IngresoDiferido"))
                Q1 = Q1 & " , CPS_CTDImputableIPE = " & vFldDao(Rs("CPS_CTDImputableIPE"))
                Q1 = Q1 & " , CPS_IncentivoAhorro = " & vFldDao(Rs("CPS_IncentivoAhorro"))
                Q1 = Q1 & " , CPS_IDPCVoluntario = " & vFldDao(Rs("CPS_IDPCVoluntario"))
                Q1 = Q1 & " , CPS_CredActFijos = " & vFldDao(Rs("CPS_CredActFijos"))
                Q1 = Q1 & " , CPS_CredParticipaciones = " & vFldDao(Rs("CPS_CredParticipaciones"))
                Q1 = Q1 & " WHERE idEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop
    End If
   End If
   Call CloseRs(Rs)


End Sub

Public Sub TrasDocumento(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)
    
   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'Documento' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE Documento ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = " SELECT IdDoc"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,IdCompCent"
    Q1 = Q1 & " ,IdCompPago"
    Q1 = Q1 & " ,TipoLib"
    Q1 = Q1 & " ,TipoDoc"
    Q1 = Q1 & " ,NumDoc"
    Q1 = Q1 & " ,NumDocHasta"
    Q1 = Q1 & " ,IdEntidad"
    Q1 = Q1 & " ,TipoEntidad"
    Q1 = Q1 & " ,RutEntidad"
    Q1 = Q1 & " ,NombreEntidad"
    Q1 = Q1 & " ,FEmision"
    Q1 = Q1 & " ,FVenc"
    Q1 = Q1 & " ,Descrip"
    Q1 = Q1 & " ,Estado"
    Q1 = Q1 & " ,Exento"
    Q1 = Q1 & " ,IdCuentaExento"
    Q1 = Q1 & " ,Afecto"
    Q1 = Q1 & " ,IdCuentaAfecto"
    Q1 = Q1 & " ,IVA"
    Q1 = Q1 & " ,IdCuentaIVA"
    Q1 = Q1 & " ,OtroImp"
    Q1 = Q1 & " ,IdCuentaOtroImp"
    Q1 = Q1 & " ,Total"
    Q1 = Q1 & " ,IdCuentaTotal"
    Q1 = Q1 & " ,IdUsuario"
    Q1 = Q1 & " ,FechaCreacion"
    Q1 = Q1 & " ,FEmisionOri"
    Q1 = Q1 & " ,CorrInterno"
    Q1 = Q1 & " ,SaldoDoc"
    Q1 = Q1 & " ,FExported"
    Q1 = Q1 & " ,OldIdDoc"
    Q1 = Q1 & " ,DTE"
    Q1 = Q1 & " ,PorcentRetencion"
    Q1 = Q1 & " ,TipoRetencion"
    Q1 = Q1 & " ,MovEdited"
    Q1 = Q1 & " ,OtrosVal"
    Q1 = Q1 & " ,FImporF29"
    Q1 = Q1 & " ,NumDocRef"
    Q1 = Q1 & " ,IdCtaBanco"
    Q1 = Q1 & " ,TipoRelEnt"
    Q1 = Q1 & " ,IdSucursal"
    Q1 = Q1 & " ,TotPagadoAnoAnt"
    Q1 = Q1 & " ,FImportSuc"
    Q1 = Q1 & " ,Giro"
    Q1 = Q1 & " ,FacCompraRetParcial"
    Q1 = Q1 & " ,IVAIrrecuperable"
    Q1 = Q1 & " ,DocOtrosEnAnalitico"
    Q1 = Q1 & " ,OldIdDocTmp"
    Q1 = Q1 & " ,NumFiscImpr"
    Q1 = Q1 & " ,NumInformeZ"
    Q1 = Q1 & " ,CantBoletas"
    Q1 = Q1 & " ,VentasAcumInfZ"
    Q1 = Q1 & " ,IdDocAsoc"
    Q1 = Q1 & " ,PropIVA"
    Q1 = Q1 & " ,ValIVAIrrec"
    Q1 = Q1 & " ,IVAInmueble"
    Q1 = Q1 & " ,FImpFacturacion"
    Q1 = Q1 & " ,CodSIIDTEIVAIrrec"
    Q1 = Q1 & " ,TipoDocAsoc"
    Q1 = Q1 & " ,IVAActFijo"
    Q1 = Q1 & " ,EntRelacionada"
    Q1 = Q1 & " ,NumCuotas"
    Q1 = Q1 & " ,CompraBienRaiz"
    Q1 = Q1 & " ,NumDocAsoc"
    Q1 = Q1 & " ,DTEDocAsoc"
    Q1 = Q1 & " ,IdANegCCosto"
    Q1 = Q1 & " ,UrlDTE"
    Q1 = Q1 & " ,CodCtaAfectoOld"
    Q1 = Q1 & " ,CodCtaExentoOld"
    Q1 = Q1 & " ,CodCtaTotalOld"
    Q1 = Q1 & " ,DocOtroEsCargo"
    Q1 = Q1 & " ,ValRet3Porc"
    Q1 = Q1 & " ,IdCuentaRet3Porc"
    Q1 = Q1 & " ,Tratamiento"
    Q1 = Q1 & " FROM Documento"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From Documento"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDoc"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT CentroCosto ON "
                Q1 = " INSERT INTO Documento"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdCompCent"
                Q1 = Q1 & " ,IdCompPago"
                Q1 = Q1 & " ,TipoLib"
                Q1 = Q1 & " ,TipoDoc"
                Q1 = Q1 & " ,NumDoc"
                Q1 = Q1 & " ,NumDocHasta"
                Q1 = Q1 & " ,IdEntidad"
                Q1 = Q1 & " ,TipoEntidad"
                Q1 = Q1 & " ,RutEntidad"
                Q1 = Q1 & " ,NombreEntidad"
                Q1 = Q1 & " ,FEmision"
                Q1 = Q1 & " ,FVenc"
                Q1 = Q1 & " ,Descrip"
                Q1 = Q1 & " ,Estado"
                Q1 = Q1 & " ,Exento"
                Q1 = Q1 & " ,IdCuentaExento"
                Q1 = Q1 & " ,Afecto"
                Q1 = Q1 & " ,IdCuentaAfecto"
                Q1 = Q1 & " ,IVA"
                Q1 = Q1 & " ,IdCuentaIVA"
                Q1 = Q1 & " ,OtroImp"
                Q1 = Q1 & " ,IdCuentaOtroImp"
                Q1 = Q1 & " ,Total"
                Q1 = Q1 & " ,IdCuentaTotal"
                Q1 = Q1 & " ,IdUsuario"
                Q1 = Q1 & " ,FechaCreacion"
                Q1 = Q1 & " ,FEmisionOri"
                Q1 = Q1 & " ,CorrInterno"
                Q1 = Q1 & " ,SaldoDoc"
                Q1 = Q1 & " ,FExported"
                Q1 = Q1 & " ,OldIdDoc"
                Q1 = Q1 & " ,DTE"
                Q1 = Q1 & " ,PorcentRetencion"
                Q1 = Q1 & " ,TipoRetencion"
                Q1 = Q1 & " ,MovEdited"
                Q1 = Q1 & " ,OtrosVal"
                Q1 = Q1 & " ,FImporF29"
                Q1 = Q1 & " ,NumDocRef"
                Q1 = Q1 & " ,IdCtaBanco"
                Q1 = Q1 & " ,TipoRelEnt"
                Q1 = Q1 & " ,IdSucursal"
                Q1 = Q1 & " ,TotPagadoAnoAnt"
                Q1 = Q1 & " ,FImportSuc"
                Q1 = Q1 & " ,Giro"
                Q1 = Q1 & " ,FacCompraRetParcial"
                Q1 = Q1 & " ,IVAIrrecuperable"
                Q1 = Q1 & " ,DocOtrosEnAnalitico"
                Q1 = Q1 & " ,OldIdDocTmp"
                Q1 = Q1 & " ,NumFiscImpr"
                Q1 = Q1 & " ,NumInformeZ"
                Q1 = Q1 & " ,CantBoletas"
                Q1 = Q1 & " ,VentasAcumInfZ"
                Q1 = Q1 & " ,IdDocAsoc"
                Q1 = Q1 & " ,PropIVA"
                Q1 = Q1 & " ,ValIVAIrrec"
                Q1 = Q1 & " ,IVAInmueble"
                Q1 = Q1 & " ,FImpFacturacion"
                Q1 = Q1 & " ,CodSIIDTEIVAIrrec"
                Q1 = Q1 & " ,TipoDocAsoc"
                Q1 = Q1 & " ,IVAActFijo"
                Q1 = Q1 & " ,EntRelacionada"
                Q1 = Q1 & " ,NumCuotas"
                Q1 = Q1 & " ,CompraBienRaiz"
                Q1 = Q1 & " ,NumDocAsoc"
                Q1 = Q1 & " ,DTEDocAsoc"
                Q1 = Q1 & " ,IdANegCCosto"
                Q1 = Q1 & " ,UrlDTE"
                Q1 = Q1 & " ,CodCtaAfectoOld"
                Q1 = Q1 & " ,CodCtaExentoOld"
                Q1 = Q1 & " ,CodCtaTotalOld"
                Q1 = Q1 & " ,DocOtroEsCargo"
                Q1 = Q1 & " ,ValRet3Porc"
                Q1 = Q1 & " ,IdCuentaRet3Porc"
                Q1 = Q1 & " ,Tratamiento"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompCent"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDoc")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDocHasta")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoEntidad"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("RutEntidad")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombreEntidad")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("FEmision"))
                Q1 = Q1 & " ," & vFldDao(Rs("FVenc"))
                Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Descrip")), Chr(39), "") & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ," & vFldDao(Rs("Exento"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaExento"))
                Q1 = Q1 & " ," & vFldDao(Rs("Afecto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaAfecto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaIVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("OtroImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaOtroImp"))
                Q1 = Q1 & " ," & vFldDao(Rs("Total"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaTotal"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaCreacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("FEmisionOri"))
                Q1 = Q1 & " ," & vFldDao(Rs("CorrInterno"))
                Q1 = Q1 & " ," & vFldDao(Rs("SaldoDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("FExported"))
                Q1 = Q1 & " ," & vFldDao(Rs("OldIdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("DTE"))
                Q1 = Q1 & " ," & vFldDao(Rs("PorcentRetencion"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoRetencion"))
                Q1 = Q1 & " ," & vFldDao(Rs("MovEdited"))
                Q1 = Q1 & " ," & vFldDao(Rs("OtrosVal"))
                Q1 = Q1 & " ," & vFldDao(Rs("FImporF29"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDocRef")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdCtaBanco"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoRelEnt"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdSucursal"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotPagadoAnoAnt"))
                Q1 = Q1 & " ," & vFldDao(Rs("FImportSuc"))
                Q1 = Q1 & " ," & vFldDao(Rs("Giro"))
                Q1 = Q1 & " ," & vFldDao(Rs("FacCompraRetParcial"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVAIrrecuperable"))
                Q1 = Q1 & " ," & vFldDao(Rs("DocOtrosEnAnalitico"))
                Q1 = Q1 & " ," & vFldDao(Rs("OldIdDocTmp"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumFiscImpr")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumInformeZ")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("CantBoletas"))
                Q1 = Q1 & " ," & vFldDao(Rs("VentasAcumInfZ"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDocAsoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("PropIVA"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValIVAIrrec"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVAInmueble"))
                Q1 = Q1 & " ," & vFldDao(Rs("FImpFacturacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("CodSIIDTEIVAIrrec"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoDocAsoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVAActFijo"))
                Q1 = Q1 & " ," & vFldDao(Rs("EntRelacionada"))
                Q1 = Q1 & " ," & vFldDao(Rs("NumCuotas"))
                Q1 = Q1 & " ," & vFldDao(Rs("CompraBienRaiz"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NumDocAsoc")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("DTEDocAsoc"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("IdANegCCosto")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("UrlDTE")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaAfectoOld")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaExentoOld")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCtaTotalOld")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("DocOtroEsCargo"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValRet3Porc"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuentaRet3Porc"))
                Q1 = Q1 & " ," & vFldDao(Rs("Tratamiento"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT CentroCosto OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE Documento"
                Q1 = Q1 & " SET IdCompCent = " & vFldDao(Rs("IdCompCent"))
                Q1 = Q1 & " ,IdCompPago = " & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ,TipoLib = " & vFldDao(Rs("TipoLib"))
                Q1 = Q1 & " ,TipoDoc = " & vFldDao(Rs("TipoDoc"))
                Q1 = Q1 & " ,NumDoc = '" & vFldDao(Rs("NumDoc")) & "'"
                Q1 = Q1 & " ,NumDocHasta = '" & vFldDao(Rs("NumDocHasta")) & "'"
                Q1 = Q1 & " ,IdEntidad = " & vFldDao(Rs("IdEntidad"))
                Q1 = Q1 & " ,TipoEntidad = " & vFldDao(Rs("TipoEntidad"))
                Q1 = Q1 & " ,RutEntidad = '" & vFldDao(Rs("RutEntidad")) & "'"
                Q1 = Q1 & " ,NombreEntidad = '" & vFldDao(Rs("NombreEntidad")) & "'"
                Q1 = Q1 & " ,FEmision = " & vFldDao(Rs("FEmision"))
                Q1 = Q1 & " ,FVenc = " & vFldDao(Rs("FVenc"))
                Q1 = Q1 & " ,Descrip = '" & Replace(vFldDao(Rs("Descrip")), Chr(39), "") & "'"
                Q1 = Q1 & " ,Estado = " & vFldDao(Rs("Estado"))
                Q1 = Q1 & " ,Exento = " & vFldDao(Rs("Exento"))
                Q1 = Q1 & " ,IdCuentaExento = " & vFldDao(Rs("IdCuentaExento"))
                Q1 = Q1 & " ,Afecto = " & vFldDao(Rs("Afecto"))
                Q1 = Q1 & " ,IdCuentaAfecto = " & vFldDao(Rs("IdCuentaAfecto"))
                Q1 = Q1 & " ,IVA = " & vFldDao(Rs("IVA"))
                Q1 = Q1 & " ,IdCuentaIVA = " & vFldDao(Rs("IdCuentaIVA"))
                Q1 = Q1 & " ,OtroImp = " & vFldDao(Rs("OtroImp"))
                Q1 = Q1 & " ,IdCuentaOtroImp = " & vFldDao(Rs("IdCuentaOtroImp"))
                Q1 = Q1 & " ,Total = " & vFldDao(Rs("Total"))
                Q1 = Q1 & " ,IdCuentaTotal = " & vFldDao(Rs("IdCuentaTotal"))
                Q1 = Q1 & " ,IdUsuario = " & vFldDao(Rs("IdUsuario"))
                Q1 = Q1 & " ,FechaCreacion = " & vFldDao(Rs("FechaCreacion"))
                Q1 = Q1 & " ,FEmisionOri = " & vFldDao(Rs("FEmisionOri"))
                Q1 = Q1 & " ,CorrInterno = " & vFldDao(Rs("CorrInterno"))
                Q1 = Q1 & " ,SaldoDoc = " & vFldDao(Rs("SaldoDoc"))
                Q1 = Q1 & " ,FExported = " & vFldDao(Rs("FExported"))
                Q1 = Q1 & " ,OldIdDoc = " & vFldDao(Rs("OldIdDoc"))
                Q1 = Q1 & " ,DTE = " & vFldDao(Rs("DTE"))
                Q1 = Q1 & " ,PorcentRetencion = " & vFldDao(Rs("PorcentRetencion"))
                Q1 = Q1 & " ,TipoRetencion = " & vFldDao(Rs("TipoRetencion"))
                Q1 = Q1 & " ,MovEdited = " & vFldDao(Rs("MovEdited"))
                Q1 = Q1 & " ,OtrosVal = " & vFldDao(Rs("OtrosVal"))
                Q1 = Q1 & " ,FImporF29 = " & vFldDao(Rs("FImporF29"))
                Q1 = Q1 & " ,NumDocRef = '" & vFldDao(Rs("NumDocRef")) & "'"
                Q1 = Q1 & " ,IdCtaBanco = " & vFldDao(Rs("IdCtaBanco"))
                Q1 = Q1 & " ,TipoRelEnt = " & vFldDao(Rs("TipoRelEnt"))
                Q1 = Q1 & " ,IdSucursal = " & vFldDao(Rs("IdSucursal"))
                Q1 = Q1 & " ,TotPagadoAnoAnt = " & vFldDao(Rs("TotPagadoAnoAnt"))
                Q1 = Q1 & " ,FImportSuc = " & vFldDao(Rs("FImportSuc"))
                Q1 = Q1 & " ,Giro = " & vFldDao(Rs("Giro"))
                Q1 = Q1 & " ,FacCompraRetParcial = " & vFldDao(Rs("FacCompraRetParcial"))
                Q1 = Q1 & " ,IVAIrrecuperable = " & vFldDao(Rs("IVAIrrecuperable"))
                Q1 = Q1 & " ,DocOtrosEnAnalitico = " & vFldDao(Rs("DocOtrosEnAnalitico"))
                Q1 = Q1 & " ,OldIdDocTmp = " & vFldDao(Rs("OldIdDocTmp"))
                Q1 = Q1 & " ,NumFiscImpr = '" & vFldDao(Rs("NumFiscImpr")) & "'"
                Q1 = Q1 & " ,NumInformeZ = '" & vFldDao(Rs("NumInformeZ")) & "'"
                Q1 = Q1 & " ,CantBoletas = " & vFldDao(Rs("CantBoletas"))
                Q1 = Q1 & " ,VentasAcumInfZ = " & vFldDao(Rs("VentasAcumInfZ"))
                Q1 = Q1 & " ,IdDocAsoc = " & vFldDao(Rs("IdDocAsoc"))
                Q1 = Q1 & " ,PropIVA = " & vFldDao(Rs("PropIVA"))
                Q1 = Q1 & " ,ValIVAIrrec = " & vFldDao(Rs("ValIVAIrrec"))
                Q1 = Q1 & " ,IVAInmueble = " & vFldDao(Rs("IVAInmueble"))
                Q1 = Q1 & " ,FImpFacturacion = " & vFldDao(Rs("FImpFacturacion"))
                Q1 = Q1 & " ,CodSIIDTEIVAIrrec = " & vFldDao(Rs("CodSIIDTEIVAIrrec"))
                Q1 = Q1 & " ,TipoDocAsoc = " & vFldDao(Rs("TipoDocAsoc"))
                Q1 = Q1 & " ,IVAActFijo = " & vFldDao(Rs("IVAActFijo"))
                Q1 = Q1 & " ,EntRelacionada = " & vFldDao(Rs("EntRelacionada"))
                Q1 = Q1 & " ,NumCuotas = " & vFldDao(Rs("NumCuotas"))
                Q1 = Q1 & " ,CompraBienRaiz = " & vFldDao(Rs("CompraBienRaiz"))
                Q1 = Q1 & " ,NumDocAsoc = '" & vFldDao(Rs("NumDocAsoc")) & "'"
                Q1 = Q1 & " ,DTEDocAsoc = " & vFldDao(Rs("DTEDocAsoc"))
                Q1 = Q1 & " ,IdANegCCosto = '" & vFldDao(Rs("IdANegCCosto")) & "'"
                Q1 = Q1 & " ,UrlDTE = '" & vFldDao(Rs("UrlDTE")) & "'"
                Q1 = Q1 & " ,CodCtaAfectoOld = '" & vFldDao(Rs("CodCtaAfectoOld")) & "'"
                Q1 = Q1 & " ,CodCtaExentoOld = '" & vFldDao(Rs("CodCtaExentoOld")) & "'"
                Q1 = Q1 & " ,CodCtaTotalOld = '" & vFldDao(Rs("CodCtaTotalOld")) & "'"
                Q1 = Q1 & " ,DocOtroEsCargo = " & vFldDao(Rs("DocOtroEsCargo"))
                Q1 = Q1 & " ,ValRet3Porc = " & vFldDao(Rs("ValRet3Porc"))
                Q1 = Q1 & " ,IdCuentaRet3Porc = " & vFldDao(Rs("IdCuentaRet3Porc"))
                Q1 = Q1 & " ,Tratamiento = " & vFldDao(Rs("Tratamiento"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdDoc"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
   
    Q1 = "UPDATE Doc"
    Q1 = Q1 & " SET    Doc.IdEntidad = IIF(EN.IdEntidad IS NULL,Doc.IdEntidad,EN.IdEntidad),"
    Q1 = Q1 & "        Doc.IdCuentaExento = IIF(CuentaEx.idCuenta IS NULL,Doc.IdCuentaExento,CuentaEx.idCuenta),"
    Q1 = Q1 & "        Doc.IdCuentaAfecto = IIF(CuentaAf.idCuenta IS NULL,Doc.IdCuentaAfecto,CuentaAf.idCuenta),"
    Q1 = Q1 & "        Doc.IdCuentaIVA = IIF(CuentaIVA.idCuenta IS NULL,Doc.IdCuentaIVA,CuentaIVA.idCuenta),"
    Q1 = Q1 & "        Doc.IdCuentaOtroImp = IIF(CuentaOImp.idCuenta IS NULL,Doc.IdCuentaOtroImp,CuentaOImp.idCuenta),"
    Q1 = Q1 & "        Doc.IdCuentaTotal = IIF(CuentaTotal.idCuenta IS NULL,Doc.IdCuentaTotal,CuentaTotal.idCuenta),"
    Q1 = Q1 & "        Doc.IdCtaBanco = IIF(CuentaBanco.idCuenta IS NULL,Doc.IdCtaBanco,CuentaBanco.idCuenta),"
    Q1 = Q1 & "        Doc.IdCuentaRet3Porc = IIF(CuentaRet3.idCuenta IS NULL,Doc.IdCuentaRet3Porc,CuentaRet3.idCuenta),"
    Q1 = Q1 & "        Doc.IdCompCent = IIF(ComCent.IdComp IS NULL,Doc.IdCompCent,ComCent.IdComp),"
    Q1 = Q1 & "        Doc.IdCompPago = IIF(ComPago.IdComp IS NULL,Doc.IdCompPago,ComPago.IdComp),"
    Q1 = Q1 & "        Doc.OldIdDocTmp = IIF(DocOld.IdDoc IS NULL,Doc.OldIdDocTmp,DocOld.IdDoc),"
    Q1 = Q1 & "        Doc.IdDocAsoc = IIF(DocAso.IdDoc IS NULL,Doc.IdDocAsoc,DocAso.IdDoc),"
    Q1 = Q1 & "        Doc.IdUsuario = IIF(Usu.IdUsuario IS NULL,Doc.IdUsuario,Usu.IdUsuario),"
    Q1 = Q1 & "        Doc.IdSucursal = IIf(Sucu.IdSucursal Is Null, Doc.IdSucursal, Sucu.IdSucursal)"
    Q1 = Q1 & " FROM ((((((((((((((Documento AS Doc LEFT JOIN Entidades AS EN ON Doc.IdEntidad = EN.idtras AND EN.IdEmpresa = Doc.IdEmpresa)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaEx ON Doc.IdCuentaExento = CuentaEx.idtras AND CuentaEx.IdEmpresa = Doc.IdEmpresa AND CuentaEx.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaAf ON Doc.IdCuentaAfecto = CuentaAf.idtras AND CuentaAf.IdEmpresa = Doc.IdEmpresa AND CuentaAf.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaIVA ON Doc.IdCuentaIVA = CuentaIVA.idtras AND CuentaIVA.IdEmpresa = Doc.IdEmpresa AND CuentaIVA.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaOImp ON Doc.IdCuentaOtroImp = CuentaOImp.idtras AND CuentaOImp.IdEmpresa = Doc.IdEmpresa AND CuentaOImp.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaTotal ON Doc.IdCuentaTotal = CuentaTotal.idtras AND CuentaTotal.IdEmpresa = Doc.IdEmpresa AND CuentaTotal.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaBanco ON Doc.IdCtaBanco = CuentaBanco.idtras AND CuentaBanco.IdEmpresa = Doc.IdEmpresa AND CuentaBanco.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas AS CuentaRet3 ON Doc.IdCuentaRet3Porc = CuentaRet3.idtras AND CuentaRet3.IdEmpresa = Doc.IdEmpresa AND CuentaRet3.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Comprobante AS ComCent ON Doc.IdCompCent = ComCent.idtras AND ComCent.IdEmpresa = Doc.IdEmpresa AND ComCent.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Comprobante AS ComPago ON Doc.IdCompPago = ComPago.idtras AND ComPago.IdEmpresa = Doc.IdEmpresa AND ComPago.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Documento AS DocOld ON Doc.OldIdDocTmp = DocOld.idtras AND DocOld.IdEmpresa = Doc.IdEmpresa AND DocOld.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Documento AS DocAso ON Doc.IdDocAsoc = DocAso.idtras AND DocAso.IdEmpresa = Doc.IdEmpresa AND DocAso.Ano = Doc.Ano)"
    Q1 = Q1 & " LEFT JOIN Usuarios AS Usu ON Doc.IdUsuario = Usu.IdUsuario)"
    Q1 = Q1 & " LEFT JOIN Sucursales AS Sucu ON Doc.IdSucursal = Sucu.IdSucursal AND Sucu.IdEmpresa = Doc.IdEmpresa)"
    Q1 = Q1 & " WHERE Doc.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND Doc.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasMovDocumento(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'MovDocumento' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE MovDocumento ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)
   

    Q1 = " SELECT IdMovDoc"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,IdDoc"
    Q1 = Q1 & " ,IdCompCent"
    Q1 = Q1 & " ,IdCompPago"
    Q1 = Q1 & " ,Orden"
    Q1 = Q1 & " ,IdCuenta"
    Q1 = Q1 & " ,Debe"
    Q1 = Q1 & " ,Haber"
    Q1 = Q1 & " ,Glosa"
    Q1 = Q1 & " ,IdTipoValLib"
    Q1 = Q1 & " ,EsTotalDoc"
    Q1 = Q1 & " ,IdCCosto"
    Q1 = Q1 & " ,IdAreaNeg"
    Q1 = Q1 & " ,Tasa"
    Q1 = Q1 & " ,EsRecuperable"
    Q1 = Q1 & " ,CodSIIDTE"
    Q1 = Q1 & " ,CodCuentaOld"
    Q1 = Q1 & " FROM MovDocumento"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From MovDocumento"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMovDoc"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT MovDocumento ON "
                Q1 = " INSERT INTO MovDocumento"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdDoc"
                Q1 = Q1 & " ,IdCompCent"
                Q1 = Q1 & " ,IdCompPago"
                Q1 = Q1 & " ,Orden"
                Q1 = Q1 & " ,IdCuenta"
                Q1 = Q1 & " ,Debe"
                Q1 = Q1 & " ,Haber"
                Q1 = Q1 & " ,Glosa"
                Q1 = Q1 & " ,IdTipoValLib"
                Q1 = Q1 & " ,EsTotalDoc"
                Q1 = Q1 & " ,IdCCosto"
                Q1 = Q1 & " ,IdAreaNeg"
                Q1 = Q1 & " ,Tasa"
                Q1 = Q1 & " ,EsRecuperable"
                Q1 = Q1 & " ,CodSIIDTE"
                Q1 = Q1 & " ,CodCuentaOld"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompCent"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ," & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ," & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ," & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,'" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdTipoValLib"))
                Q1 = Q1 & " ," & vFldDao(Rs("EsTotalDoc"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCCosto"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAreaNeg"))
                Q1 = Q1 & " ," & Replace(vFldDao(Rs("Tasa")), ",", ".")
                Q1 = Q1 & " ," & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodSIIDTE")) & "'"
                Q1 = Q1 & " ,'" & vFldDao(Rs("CodCuentaOld")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdMovDoc")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT MovDocumento OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE MovDocumento"
                Q1 = Q1 & " SET IdDoc = " & vFldDao(Rs("IdDoc"))
                Q1 = Q1 & " ,IdCompCent = " & vFldDao(Rs("IdCompCent"))
                Q1 = Q1 & " ,IdCompPago = " & vFldDao(Rs("IdCompPago"))
                Q1 = Q1 & " ,Orden = " & vFldDao(Rs("Orden"))
                Q1 = Q1 & " ,IdCuenta = " & vFldDao(Rs("IdCuenta"))
                Q1 = Q1 & " ,Debe = " & vFldDao(Rs("Debe"))
                Q1 = Q1 & " ,Haber = " & vFldDao(Rs("Haber"))
                Q1 = Q1 & " ,Glosa = '" & Replace(vFldDao(Rs("Glosa")), Chr(39), "") & "'"
                Q1 = Q1 & " ,IdTipoValLib = " & vFldDao(Rs("IdTipoValLib"))
                Q1 = Q1 & " ,EsTotalDoc = " & vFldDao(Rs("EsTotalDoc"))
                Q1 = Q1 & " ,IdCCosto = " & vFldDao(Rs("IdCCosto"))
                Q1 = Q1 & " ,IdAreaNeg = " & vFldDao(Rs("IdAreaNeg"))
                Q1 = Q1 & " ,Tasa = " & Replace(vFldDao(Rs("Tasa")), ",", ".")
                Q1 = Q1 & " ,EsRecuperable = " & vFldDao(Rs("EsRecuperable"))
                Q1 = Q1 & " ,CodSIIDTE = '" & vFldDao(Rs("CodSIIDTE")) & "'"
                Q1 = Q1 & " ,CodCuentaOld = '" & vFldDao(Rs("CodCuentaOld")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdMovDoc"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE MD "
    Q1 = Q1 & " SET    MD.IdDoc = IIF(DOC.IdDoc IS NULL,0,DOC.IdDoc),"
    Q1 = Q1 & "        MD.IdCompCent = IIF(COMC.IdComp IS NULL,MD.IdCompCent,COMC.IdComp),"
    Q1 = Q1 & "        MD.IdCompPago = IIF(COMP.IdComp IS NULL,MD.IdCompPago,COMP.IdComp),"
    Q1 = Q1 & "        MD.IdCuenta = IIF(CU.idCuenta IS NULL,MD.IdCuenta,CU.idCuenta),"
    Q1 = Q1 & "        MD.IdCCosto = IIF(CC.IdCCosto IS NULL,MD.IdCCosto,CC.IdCCosto),"
    Q1 = Q1 & "        MD.IdAreaNeg = IIf(AN.IdAreaNegocio Is Null, MD.IdAreaNeg, AN.IdAreaNegocio)"
    Q1 = Q1 & " FROM ((((((MovDocumento MD"
    Q1 = Q1 & " LEFT JOIN Documento DOC ON DOC.IdTras = MD.IdDoc AND DOC.IdEmpresa = MD.IdEmpresa AND DOC.Ano = MD.Ano)"
    Q1 = Q1 & " LEFT JOIN Comprobante COMC ON COMC.IdTras = MD.IdCompCent AND COMC.IdEmpresa = MD.IdEmpresa AND COMC.Ano = MD.Ano)"
    Q1 = Q1 & " LEFT JOIN Comprobante COMP ON COMP.IdTras = MD.IdCompPago AND COMP.IdEmpresa = MD.IdEmpresa AND COMP.Ano = MD.Ano)"
    Q1 = Q1 & " LEFT JOIN Cuentas CU ON CU.IdTras = MD.IdCuenta AND CU.IdEmpresa = MD.IdEmpresa AND CU.Ano = MD.Ano)"
    Q1 = Q1 & " LEFT JOIN CentroCosto CC ON CC.IdTras = MD.IdCCosto AND CC.IdEmpresa = MD.IdEmpresa)"
    Q1 = Q1 & " LEFT JOIN AreaNegocio AN ON AN.IdTras = MD.IdAreaNeg AND AN.IdEmpresa = MD.IdEmpresa)"
    Q1 = Q1 & " WHERE MD.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND MD.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub



Public Sub TrasActFijoCompsFicha(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'ActFijoCompsFicha' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE ActFijoCompsFicha ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

   Q1 = "SELECT IdCompFicha "
   Q1 = Q1 & " ,IdEmpresa "
   Q1 = Q1 & " ,Ano "
   Q1 = Q1 & " ,IdActFijo "
   Q1 = Q1 & " ,IdGrupo "
   Q1 = Q1 & " ,IdComp "
   Q1 = Q1 & " ,PjeDivComp "
   Q1 = Q1 & " ,ValorCompra "
   Q1 = Q1 & " ,ValorResidual "
   Q1 = Q1 & " ,PjeAmortizacion "
   Q1 = Q1 & " ,VidaUtil "
   Q1 = Q1 & " ,CostosAdicionales "
   Q1 = Q1 & " ,TasaDesc "
   Q1 = Q1 & " ,CostoDesmant "
   Q1 = Q1 & " ,ValActCostoDesmant "
   Q1 = Q1 & " ,ValorBien "
   Q1 = Q1 & " ,ValorRazonable_31_12 "
   Q1 = Q1 & " ,NoExisteValRazonable "
   Q1 = Q1 & " ,OtrasDiferencias "
   Q1 = Q1 & " ,DepAcum "
   Q1 = Q1 & " ,VidaUtilDep "
   Q1 = Q1 & " ,ReservaAcum "
   Q1 = Q1 & " ,DepAcumuladaAnoAnt "
   Q1 = Q1 & " ,VidaUtilYaDep "
   Q1 = Q1 & " ,ReservaAcumAnt "
   Q1 = Q1 & " ,IdCompFichaOldTmp "
   Q1 = Q1 & " ,IdCompFichaOld "
   Q1 = Q1 & " ,DepPeriodo "
   Q1 = Q1 & " ,Factor "
   Q1 = Q1 & " ,Revalorizacion "
   Q1 = Q1 & " From ActFijoCompsFicha "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
   Q1 = Q1 & " AND Ano = " & Ano
   Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From ActFijoCompsFicha"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCompFicha"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT ActFijoCompsFicha ON "
                Q1 = " INSERT INTO ActFijoCompsFicha "
                Q1 = Q1 & " (IdEmpresa "
                Q1 = Q1 & " ,Ano "
                Q1 = Q1 & " ,IdActFijo "
                Q1 = Q1 & " ,IdGrupo "
                Q1 = Q1 & " ,IdComp "
                Q1 = Q1 & " ,PjeDivComp "
                Q1 = Q1 & " ,ValorCompra "
                Q1 = Q1 & " ,ValorResidual "
                Q1 = Q1 & " ,PjeAmortizacion "
                Q1 = Q1 & " ,VidaUtil "
                Q1 = Q1 & " ,CostosAdicionales "
                Q1 = Q1 & " ,TasaDesc "
                Q1 = Q1 & " ,CostoDesmant "
                Q1 = Q1 & " ,ValActCostoDesmant "
                Q1 = Q1 & " ,ValorBien "
                Q1 = Q1 & " ,ValorRazonable_31_12 "
                Q1 = Q1 & " ,NoExisteValRazonable "
                Q1 = Q1 & " ,OtrasDiferencias "
                Q1 = Q1 & " ,DepAcum "
                Q1 = Q1 & " ,VidaUtilDep "
                Q1 = Q1 & " ,ReservaAcum "
                Q1 = Q1 & " ,DepAcumuladaAnoAnt "
                Q1 = Q1 & " ,VidaUtilYaDep "
                Q1 = Q1 & " ,ReservaAcumAnt "
                Q1 = Q1 & " ,IdCompFichaOldTmp "
                Q1 = Q1 & " ,IdCompFichaOld "
                Q1 = Q1 & " ,DepPeriodo "
                Q1 = Q1 & " ,Factor "
                Q1 = Q1 & " ,Revalorizacion "
                Q1 = Q1 & " ,IdTras) "
                Q1 = Q1 & "  Values "
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdActFijo"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("PjeDivComp"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValorCompra"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValorResidual"))
                Q1 = Q1 & " ," & vFldDao(Rs("PjeAmortizacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtil"))
                Q1 = Q1 & " ," & vFldDao(Rs("CostosAdicionales"))
                Q1 = Q1 & " ," & vFldDao(Rs("TasaDesc"))
                Q1 = Q1 & " ," & vFldDao(Rs("CostoDesmant"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValActCostoDesmant"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValorBien"))
                Q1 = Q1 & " ," & vFldDao(Rs("ValorRazonable_31_12"))
                Q1 = Q1 & " ," & vFldDao(Rs("NoExisteValRazonable"))
                Q1 = Q1 & " ," & vFldDao(Rs("OtrasDiferencias"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAcum"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtilDep"))
                Q1 = Q1 & " ," & vFldDao(Rs("ReservaAcum"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepAcumuladaAnoAnt"))
                Q1 = Q1 & " ," & vFldDao(Rs("VidaUtilYaDep"))
                Q1 = Q1 & " ," & vFldDao(Rs("ReservaAcumAnt"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompFichaOldTmp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompFichaOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("DepPeriodo"))
                Q1 = Q1 & " ," & vFldDao(Rs("Factor"))
                Q1 = Q1 & " ," & vFldDao(Rs("Revalorizacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdCompFicha")) & ") "
                'Q1 = Q1 & " SET IDENTITY_INSERT ActFijoCompsFicha OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE ActFijoCompsFicha"
                Q1 = Q1 & " SET IdActFijo = " & vFldDao(Rs("IdActFijo"))
                Q1 = Q1 & " ,IdGrupo = " & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ,IdComp = " & vFldDao(Rs("IdComp"))
                Q1 = Q1 & " ,PjeDivComp = " & vFldDao(Rs("PjeDivComp"))
                Q1 = Q1 & " ,ValorCompra = " & vFldDao(Rs("ValorCompra"))
                Q1 = Q1 & " ,ValorResidual = " & vFldDao(Rs("ValorResidual"))
                Q1 = Q1 & " ,PjeAmortizacion = " & vFldDao(Rs("PjeAmortizacion"))
                Q1 = Q1 & " ,VidaUtil = " & vFldDao(Rs("VidaUtil"))
                Q1 = Q1 & " ,CostosAdicionales = " & vFldDao(Rs("CostosAdicionales"))
                Q1 = Q1 & " ,TasaDesc = " & vFldDao(Rs("TasaDesc"))
                Q1 = Q1 & " ,CostoDesmant = " & vFldDao(Rs("CostoDesmant"))
                Q1 = Q1 & " ,ValActCostoDesmant = " & vFldDao(Rs("ValActCostoDesmant"))
                Q1 = Q1 & " ,ValorBien = " & vFldDao(Rs("ValorBien"))
                Q1 = Q1 & " ,ValorRazonable_31_12 = " & vFldDao(Rs("ValorRazonable_31_12"))
                Q1 = Q1 & " ,NoExisteValRazonable = " & vFldDao(Rs("NoExisteValRazonable"))
                Q1 = Q1 & " ,OtrasDiferencias = " & vFldDao(Rs("OtrasDiferencias"))
                Q1 = Q1 & " ,DepAcum = " & vFldDao(Rs("DepAcum"))
                Q1 = Q1 & " ,VidaUtilDep = " & vFldDao(Rs("VidaUtilDep"))
                Q1 = Q1 & " ,ReservaAcum = " & vFldDao(Rs("ReservaAcum"))
                Q1 = Q1 & " ,DepAcumuladaAnoAnt = " & vFldDao(Rs("DepAcumuladaAnoAnt"))
                Q1 = Q1 & " ,VidaUtilYaDep = " & vFldDao(Rs("VidaUtilYaDep"))
                Q1 = Q1 & " ,ReservaAcumAnt = " & vFldDao(Rs("ReservaAcumAnt"))
                Q1 = Q1 & " ,IdCompFichaOldTmp = " & vFldDao(Rs("IdCompFichaOldTmp"))
                Q1 = Q1 & " ,IdCompFichaOld = " & vFldDao(Rs("IdCompFichaOld"))
                Q1 = Q1 & " ,DepPeriodo = " & vFldDao(Rs("DepPeriodo"))
                Q1 = Q1 & " ,Factor = " & vFldDao(Rs("Factor"))
                Q1 = Q1 & " ,Revalorizacion = " & vFldDao(Rs("Revalorizacion"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdCompFicha"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = " UPDATE AFC"
    Q1 = Q1 & " SET    AFC.IdActFijo = ISNULL(MA.IdActFijo,AFC.IdActFijo),"
    Q1 = Q1 & "        AFC.IdGrupo = ISNULL(AG.IdGrupo,AFC.IdGrupo),"
    Q1 = Q1 & "        AFC.IdComp = ISNULL(AFCO.IdComp,AFC.IdComp),"
    Q1 = Q1 & "        AFC.IdCompFichaOld = ISNULL(AFCOM.IdCompFicha,AFC.IdCompFichaOld),"
    Q1 = Q1 & "        AFC.IdCompFichaOldTmp = ISNULL(AFCOMP.IdCompFicha, AFC.IdCompFichaOldTmp)"
    Q1 = Q1 & " FROM ActFijoCompsFicha AFC"
    Q1 = Q1 & " LEFT JOIN MovActivoFijo MA ON MA.IdTras = AFC.IdActFijo AND MA.IdEmpresa = AFC.IdEmpresa AND MA.Ano = AFC.Ano"
    Q1 = Q1 & " LEFT JOIN AFGrupos AG ON AG.IdTras = AFC.IdGrupo AND AG.IdEmpresa = AFC.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN AFComponentes AFCO ON AFCO.IdTras = AFC.IdComp AND AFCO.IdEmpresa = AFC.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN ActFijoCompsFicha AFCOM ON AFCOM.IdTras = AFC.IdCompFichaOld AND AFCOM.IdEmpresa = AFC.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN ActFijoCompsFicha AFCOMP ON AFCOMP.IdTras = AFC.IdCompFichaOldTmp AND AFCOMP.IdEmpresa = AFC.IdEmpresa"
    Q1 = Q1 & " WHERE AFC.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   AFC.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)


End Sub

Public Sub TrasActFijoFicha(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'ActFijoFicha' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE ActFijoFicha ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)
    
    Q1 = "SELECT IdFicha "
    Q1 = Q1 & " ,IdEmpresa "
    Q1 = Q1 & " ,Ano "
    Q1 = Q1 & " ,IdActFijo "
    Q1 = Q1 & " ,IdGrupo "
    Q1 = Q1 & " ,PrecioFactura "
    Q1 = Q1 & " ,DerechosIntern "
    Q1 = Q1 & " ,Transporte "
    Q1 = Q1 & " ,ObrasAdapt "
    Q1 = Q1 & " ,PrecioAdquis "
    Q1 = Q1 & " ,IVARecuperable "
    Q1 = Q1 & " ,FormacionPers "
    Q1 = Q1 & " ,ObrasReubic "
    Q1 = Q1 & " ,TotalGastos "
    Q1 = Q1 & " ,FechaIncorporacion "
    Q1 = Q1 & " ,FechaDisponible "
    Q1 = Q1 & " ,AdquiOtrosConceptos "
    Q1 = Q1 & " ,GastoOtrosConceptos "
    Q1 = Q1 & " ,SinDetComps "
    Q1 = Q1 & " ,IdFichaOldTmp "
    Q1 = Q1 & " ,IdFichaOld "
    Q1 = Q1 & " From ActFijoFicha "
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From ActFijoFicha"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdFicha"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT ActFijoFicha ON "
                Q1 = " INSERT INTO ActFijoFicha"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,IdActFijo"
                Q1 = Q1 & " ,IdGrupo"
                Q1 = Q1 & " ,PrecioFactura"
                Q1 = Q1 & " ,DerechosIntern"
                Q1 = Q1 & " ,Transporte"
                Q1 = Q1 & " ,ObrasAdapt"
                Q1 = Q1 & " ,PrecioAdquis"
                Q1 = Q1 & " ,IVARecuperable"
                Q1 = Q1 & " ,FormacionPers"
                Q1 = Q1 & " ,ObrasReubic"
                Q1 = Q1 & " ,TotalGastos"
                Q1 = Q1 & " ,FechaIncorporacion"
                Q1 = Q1 & " ,FechaDisponible"
                Q1 = Q1 & " ,AdquiOtrosConceptos"
                Q1 = Q1 & " ,GastoOtrosConceptos"
                Q1 = Q1 & " ,SinDetComps"
                Q1 = Q1 & " ,IdFichaOldTmp"
                Q1 = Q1 & " ,IdFichaOld"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdActFijo"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ," & vFldDao(Rs("PrecioFactura"))
                Q1 = Q1 & " ," & vFldDao(Rs("DerechosIntern"))
                Q1 = Q1 & " ," & vFldDao(Rs("Transporte"))
                Q1 = Q1 & " ," & vFldDao(Rs("ObrasAdapt"))
                Q1 = Q1 & " ," & vFldDao(Rs("PrecioAdquis"))
                Q1 = Q1 & " ," & vFldDao(Rs("IVARecuperable"))
                Q1 = Q1 & " ," & vFldDao(Rs("FormacionPers"))
                Q1 = Q1 & " ," & vFldDao(Rs("ObrasReubic"))
                Q1 = Q1 & " ," & vFldDao(Rs("TotalGastos"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaIncorporacion"))
                Q1 = Q1 & " ," & vFldDao(Rs("FechaDisponible"))
                Q1 = Q1 & " ," & vFldDao(Rs("AdquiOtrosConceptos"))
                Q1 = Q1 & " ," & vFldDao(Rs("GastoOtrosConceptos"))
                Q1 = Q1 & " ," & vFldDao(Rs("SinDetComps"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdFichaOldTmp"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdFichaOld"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdFicha")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT ActFijoFicha OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE ActFijoFicha"
                Q1 = Q1 & " SET IdActFijo = " & vFldDao(Rs("IdActFijo"))
                Q1 = Q1 & " ,IdGrupo = " & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ,PrecioFactura = " & vFldDao(Rs("PrecioFactura"))
                Q1 = Q1 & " ,DerechosIntern = " & vFldDao(Rs("DerechosIntern"))
                Q1 = Q1 & " ,Transporte = " & vFldDao(Rs("Transporte"))
                Q1 = Q1 & " ,ObrasAdapt = " & vFldDao(Rs("ObrasAdapt"))
                Q1 = Q1 & " ,PrecioAdquis = " & vFldDao(Rs("PrecioAdquis"))
                Q1 = Q1 & " ,IVARecuperable = " & vFldDao(Rs("IVARecuperable"))
                Q1 = Q1 & " ,FormacionPers = " & vFldDao(Rs("FormacionPers"))
                Q1 = Q1 & " ,ObrasReubic = " & vFldDao(Rs("ObrasReubic"))
                Q1 = Q1 & " ,TotalGastos = " & vFldDao(Rs("TotalGastos"))
                Q1 = Q1 & " ,FechaIncorporacion = " & vFldDao(Rs("FechaIncorporacion"))
                Q1 = Q1 & " ,FechaDisponible = " & vFldDao(Rs("FechaDisponible"))
                Q1 = Q1 & " ,AdquiOtrosConceptos = " & vFldDao(Rs("AdquiOtrosConceptos"))
                Q1 = Q1 & " ,GastoOtrosConceptos = " & vFldDao(Rs("GastoOtrosConceptos"))
                Q1 = Q1 & " ,SinDetComps = " & vFldDao(Rs("SinDetComps"))
                Q1 = Q1 & " ,IdFichaOldTmp = " & vFldDao(Rs("IdFichaOldTmp"))
                Q1 = Q1 & " ,IdFichaOld = " & vFldDao(Rs("IdFichaOld"))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdFicha"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
    Q1 = "UPDATE AF"
    Q1 = Q1 & " SET  AF.IdActFijo = MA.IdActFijo,"
    Q1 = Q1 & "      AF.IdGrupo = AG.IdGrupo,"
    Q1 = Q1 & "      AF.IdFichaOld = ISNULL(ACF.IdActFijo,AF.IdFichaOld),"
    Q1 = Q1 & "      AF.IdFichaOldTmp = ISNULL(ACFI.IdActFijo, AF.IdFichaOldTmp)"
    Q1 = Q1 & " FROM ActFijoFicha AF"
    Q1 = Q1 & " LEFT JOIN MovActivoFijo MA ON MA.IdTras = AF.IdActFijo AND MA.IdEmpresa = AF.IdEmpresa AND MA.Ano = AF.Ano"
    Q1 = Q1 & " LEFT JOIN AFGrupos AG ON AG.IdTras = AF.IdGrupo AND AG.IdEmpresa = AF.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN ActFijoFicha ACF ON ACF.IdTras = AF.IdFichaOld AND ACF.IdEmpresa = AF.IdEmpresa"
    Q1 = Q1 & " LEFT JOIN ActFijoFicha ACFI ON ACFI.IdTras = AF.IdFichaOldTmp AND ACFI.IdEmpresa = AF.IdEmpresa"
    Q1 = Q1 & " WHERE AF.IdEmpresa = " & IdEmpresa
    Q1 = Q1 & " AND   AF.Ano = " & Ano
    Call ExecSQL(DBSql, Q1)
                
End Sub

Public Sub TrasAFComponentes(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'AFComponentes' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE AFComponentes ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdComp"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,IdGrupo"
    Q1 = Q1 & " ,NombComp"
    Q1 = Q1 & " From AFComponentes"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AFComponentes"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT AFComponentes ON "
                Q1 = " INSERT INTO AFComponentes"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,IdGrupo"
                Q1 = Q1 & " ,NombComp"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombComp")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdComp")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AFComponentes OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AFComponentes"
                Q1 = Q1 & " SET IdGrupo = " & vFldDao(Rs("IdGrupo"))
                Q1 = Q1 & " ,NombComp = '" & vFldDao(Rs("NombComp")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdComp"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
   
   
    Q1 = " UPDATE C"
    Q1 = Q1 & " Set c.IdGrupo = IsNull(G.IdGrupo, c.IdGrupo)"
    Q1 = Q1 & " FROM (AFComponentes C LEFT JOIN AFGrupos G ON C.IdGrupo = G.IdTras AND G.IdEmpresa = C.IdEmpresa)"
    Q1 = Q1 & " WHERE c.IdEmpresa = " & IdEmpresa
    Call ExecSQL(DBSql, Q1)
   
   
End Sub

Public Sub TrasAFGrupos(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'AFGrupos' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE AFGrupos ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdGrupo"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,NombGrupo"
    Q1 = Q1 & " From AFGrupos"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AFGrupos"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdGrupo"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                'Q1 = " SET IDENTITY_INSERT AFGrupos ON "
                Q1 = " INSERT INTO AFGrupos"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,NombGrupo"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ,'" & vFldDao(Rs("NombGrupo")) & "'"
                Q1 = Q1 & " ," & vFldDao(Rs("IdGrupo")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AFGrupos OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AFGrupos"
                Q1 = Q1 & " SET NombGrupo = '" & vFldDao(Rs("NombGrupo")) & "'"
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdGrupo"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
End Sub

Public Sub TrasAfiliadoVoluntario()

'    Q1 = "SELECT IdAfVol"
'    Q1 = Q1 & " ,IdEmpl"
'    Q1 = Q1 & " ,IdAFP"
'    Q1 = Q1 & " ,AnoMes"
'    Q1 = Q1 & " ,TipoAfVol"
'    Q1 = Q1 & " ,ValorCapVol"
'    Q1 = Q1 & " ,ValorAhorroVol"
'    Q1 = Q1 & " ,UltimaCotizacion"
'    Q1 = Q1 & " ,IdEmpresa"
'    Q1 = Q1 & " From AfiliadoVoluntario"
'
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT AfiliadoVoluntario ON "
'      Q1 = Q1 & " INSERT INTO AfiliadoVoluntario"
'      Q1 = Q1 & " (IdAfVol"
'      Q1 = Q1 & " ,IdEmpl"
'      Q1 = Q1 & " ,IdAFP"
'      Q1 = Q1 & " ,AnoMes"
'      Q1 = Q1 & " ,TipoAfVol"
'      Q1 = Q1 & " ,ValorCapVol"
'      Q1 = Q1 & " ,ValorAhorroVol"
'      Q1 = Q1 & " ,UltimaCotizacion"
'      Q1 = Q1 & " ,IdEmpresa)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("IdAfVol"))
'      Q1 = Q1 & " ," & vFld(Rs("IdEmpl"))
'      Q1 = Q1 & " ," & vFld(Rs("IdAFP"))
'      Q1 = Q1 & " ," & vFld(Rs("AnoMes"))
'      Q1 = Q1 & " ," & vFld(Rs("TipoAfVol"))
'      Q1 = Q1 & " ," & vFld(Rs("ValorCapVol"))
'      Q1 = Q1 & " ," & vFld(Rs("ValorAhorroVol"))
'      Q1 = Q1 & " ," & vFld(Rs("UltimaCotizacion"))
'      Q1 = Q1 & " ," & vFld(Rs("IdEmpresa")) & ")"
'      Q1 = Q1 & " SET IDENTITY_INSERT AfiliadoVoluntario OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)
End Sub

Public Sub TrasAFP()

'    Q1 = "SELECT idAFP"
'    Q1 = Q1 & " ,AFP"
'    Q1 = Q1 & " ,ArchPago"
'    Q1 = Q1 & " ,ArchNoPago"
'    Q1 = Q1 & " ,CodPrevired"
'    Q1 = Q1 & " From AFP"
'
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT AFP ON "
'      Q1 = Q1 & " INSERT INTO AFP"
'      Q1 = Q1 & " (idAFP"
'      Q1 = Q1 & " ,AFP"
'      Q1 = Q1 & " ,ArchPago"
'      Q1 = Q1 & " ,ArchNoPago"
'      Q1 = Q1 & " ,CodPrevired)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("idAFP"))
'      Q1 = Q1 & " ,'" & vFld(Rs("AFP")) & "'"
'      Q1 = Q1 & " ,'" & vFld(Rs("ArchPago")) & "'"
'      Q1 = Q1 & " ,'" & vFld(Rs("ArchNoPago")) & "'"
'      Q1 = Q1 & " ,'" & vFld(Rs("CodPrevired")) & "')"
'      Q1 = Q1 & " SET IDENTITY_INSERT AFP OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)
End Sub

Public Sub TrasAjusteIVAMensual(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long

    Q1 = "SELECT IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,Mes"
    Q1 = Q1 & " ,Valor"
    Q1 = Q1 & " From AjusteIVAMensual"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AjusteIVAMensual"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " INSERT INTO AjusteIVAMensual"
                Q1 = Q1 & " (IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,Mes"
                Q1 = Q1 & " ,Valor)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("Mes"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Valor")))) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AjusteIVAMensual OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AjusteIVAMensual"
                Q1 = Q1 & " SET Valor = " & str(vFmt(vFldDao(Rs("Valor"))))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND Mes = " & vFldDao(Rs("Mes"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
End Sub

Public Sub TrasAjustesExtLibCaja(DBSql As ADODB.Connection, DbAccess As Database, IdEmpresa As Long, Ano As Integer)

   Dim Rs As dao.Recordset
   Dim Rs1 As Recordset
   Dim CantSql As Long
   
    Q1 = ""
    Q1 = Q1 & "IF NOT EXISTS( "
    Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
    Q1 = Q1 & "WHERE TABLE_NAME = 'AjustesExtLibCaja' AND COLUMN_NAME = 'IdTras' "
    Q1 = Q1 & ")BEGIN "
    Q1 = Q1 & "ALTER TABLE AjustesExtLibCaja ADD IdTras INT NULL; "
    Q1 = Q1 & "END "
    
    Call ExecSQL(DBSql, Q1)

    Q1 = "SELECT IdAjustesExtLibCaja"
    Q1 = Q1 & " ,IdEmpresa"
    Q1 = Q1 & " ,Ano"
    Q1 = Q1 & " ,TipoAjuste"
    Q1 = Q1 & " ,IdItemAjuste"
    Q1 = Q1 & " ,Valor"
    Q1 = Q1 & " From AjustesExtLibCaja"
    Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresaTras
    Q1 = Q1 & " AND Ano = " & Ano
    Set Rs = OpenRsDao(DbAccess, Q1)

   
   Do While Rs.EOF = False
       
            Q1 = "SELECT * "
            Q1 = Q1 & " From AjustesExtLibCaja"
            Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
            Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
            Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAjustesExtLibCaja"))
            Set Rs1 = OpenRs(DBSql, Q1)
            
            If Rs1.EOF = True Then
            
                Q1 = " INSERT INTO AjustesExtLibCaja"
                Q1 = Q1 & " (IdAjustesExtLibCaja"
                Q1 = Q1 & " ,IdEmpresa"
                Q1 = Q1 & " ,Ano"
                Q1 = Q1 & " ,TipoAjuste"
                Q1 = Q1 & " ,IdItemAjuste"
                Q1 = Q1 & " ,Valor"
                Q1 = Q1 & " ,IdTras)"
                Q1 = Q1 & " Values"
                Q1 = Q1 & " (" & vFldDao(Rs("IdAjustesExtLibCaja"))
                Q1 = Q1 & " ," & IdEmpresa
                Q1 = Q1 & " ," & vFldDao(Rs("Ano"))
                Q1 = Q1 & " ," & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ," & vFldDao(Rs("IdItemAjuste"))
                Q1 = Q1 & " ," & str(vFmt(vFldDao(Rs("Valor"))))
                Q1 = Q1 & " ," & vFldDao(Rs("IdAjustesExtLibCaja")) & ")"
                'Q1 = Q1 & " SET IDENTITY_INSERT AjustesExtLibCaja OFF  "
                Call ExecSQL(DBSql, Q1)
                
            Else
            
                Q1 = " UPDATE AjustesExtLibCaja"
                Q1 = Q1 & " SET TipoAjuste = " & vFldDao(Rs("TipoAjuste"))
                Q1 = Q1 & " ,IdItemAjuste = " & vFldDao(Rs("IdItemAjuste"))
                Q1 = Q1 & " ,Valor = " & str(vFmt(vFldDao(Rs("Valor"))))
                Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
                Q1 = Q1 & " AND Ano = " & vFldDao(Rs("Ano"))
                Q1 = Q1 & " AND IdTras = " & vFldDao(Rs("IdAjustesExtLibCaja"))
                Call ExecSQL(DBSql, Q1)
                
            End If
            Call CloseRs(Rs1)

      Rs.MoveNext
      Loop

   Call CloseRs(Rs)
End Sub

Public Sub TrasAlertas()

'    Q1 = "SELECT idAlerta"
'    Q1 = Q1 & " ,idEmpresa"
'    Q1 = Q1 & " ,AnoMes"
'    Q1 = Q1 & " ,idEmpl"
'    Q1 = Q1 & " ,Mensaje"
'    Q1 = Q1 & " From Alertas"
'
'   Set Rs = OpenRs(DbMain, Q1)
'
'
'   Do While Rs.EOF = False
'      Q1 = " SET IDENTITY_INSERT Alertas ON "
'      Q1 = Q1 & " INSERT INTO Alertas"
'      Q1 = Q1 & " (idAlerta"
'      Q1 = Q1 & " ,idEmpresa"
'      Q1 = Q1 & " ,AnoMes"
'      Q1 = Q1 & " ,idEmpl"
'      Q1 = Q1 & " ,Mensaje)"
'      Q1 = Q1 & " Values"
'      Q1 = Q1 & " (" & vFld(Rs("idAlerta"))
'      Q1 = Q1 & " ," & vFld(Rs("idEmpresa"))
'      Q1 = Q1 & " ," & vFld(Rs("AnoMes"))
'      Q1 = Q1 & " ," & vFld(Rs("idEmpl"))
'      Q1 = Q1 & " ,'" & vFld(Rs("Mensaje")) & "')"
'      Q1 = Q1 & " SET IDENTITY_INSERT Alertas OFF  "
'      Call ExecSQL(Db, Q1)
'
'      Rs.MoveNext
'   Loop
'   Call CloseRs(Rs)
End Sub
