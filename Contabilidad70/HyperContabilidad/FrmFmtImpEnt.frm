VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFmtImpEnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato de Importación de Entidades"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "FrmFmtImpEnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_VerEjemplo 
      Caption         =   "Ver ejemplo...."
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   9120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   10275
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8760
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "FrmFmtImpEnt.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   1260
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   240
      TabIndex        =   6
      Top             =   6060
      Width           =   10335
      Begin VB.Label Lb_OtrosDocs 
         Caption         =   "Los documentos importados quedarán en estado APROBADO"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   2400
         Width           =   8955
      End
      Begin VB.Label Lb_NombCuentas 
         Caption         =   "El sistema no importa el  nombre de la cuenta, dado que se utiliza el que está en el plan de cuentas definido para la empresa"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   8955
      End
      Begin VB.Label Lb_IFRS 
         Caption         =   "Recuerde que si Ud. importa un plan de cuentas, deberá configurar manualmente las cuentas de IFRS"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   2460
         Width           =   8955
      End
      Begin VB.Label Lb_Nulos 
         Caption         =   "Documentos Anulados : RUT en blanco,  Razón Social = NULO,  Descripción = NULO,  valores en cero."
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   8955
      End
      Begin VB.Label Label1 
         Caption         =   "El formato del archivo es posicional, por lo que se deben incluir  TODOS los campos, aunque vayan en blanco. "
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   8955
      End
      Begin VB.Label Label3 
         Caption         =   "NOTAS:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lb_NotaImp 
         Caption         =   $"FrmFmtImpEnt.frx":0451
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   8955
      End
      Begin VB.Label Label1 
         Caption         =   "* Indica los campos que deben tener un valor válido (distinto de blanco)"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   1620
         Width           =   8955
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   1380
      TabIndex        =   5
      Top             =   7140
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas o campos del archivo:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2475
   End
End
Attribute VB_Name = "FrmFmtImpEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const C_CAMPO = 0
Const C_FORMATO = 1

Dim lFmtArray() As FmtImp_t
Dim lFmtCaption As String
Dim lRepCaptura As Boolean
Dim lEjemplo As Boolean
Dim lLbIFRS As Boolean
Dim lIFRS As Boolean
Dim lOtrosDocs As Boolean
Dim lNombCuentas As Boolean
Dim lNulos As Boolean


Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2ClipTrasp(Grid, Me.Caption)
End Sub

Private Sub Bt_VerEjemplo_Click()
   Dim Frm As FrmEjemploImport
   
   Set Frm = New FrmEjemploImport
   Call Frm.FViewComprobantes
   Set Frm = Nothing
   
End Sub

Private Sub Form_Load()
   
   Me.Caption = lFmtCaption
   If lRepCaptura Then
      Lb_NotaImp.Caption = ReplaceStr(Lb_NotaImp, "importación", "captura")
   End If
   
   Lb_Nulos.visible = False
   If lNulos Then
      Lb_Nulos.visible = True
   End If
   
   Lb_IFRS.visible = False
   If lIFRS Then
      Lb_IFRS.visible = True
   End If
   
   Lb_NombCuentas.visible = False
   If lNombCuentas Then
      Lb_NombCuentas.visible = True
   End If
   
   If lLbIFRS Then
      If Not Lb_Nulos.visible Then
         Lb_IFRS.Top = Lb_Nulos.Top
      End If
   End If
   
   If lNombCuentas Then
      If Lb_Nulos.visible Then
         Lb_NombCuentas.Top = Lb_IFRS.Top
      End If
   End If
   
   Lb_OtrosDocs.visible = False
   If lOtrosDocs Then
      Lb_OtrosDocs.visible = True
   End If
   
   Call SetUpGrid
   Call LoadGrid
   
   If lEjemplo Then
      Bt_VerEjemplo.visible = True
      Me.Height = Me.Height + Bt_VerEjemplo.Height + 200
   End If


End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)

   Grid.ColWidth(C_CAMPO) = 2400
   Grid.ColWidth(C_FORMATO) = 8400
   
   Grid.ColAlignment(C_CAMPO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FORMATO) = flexAlignLeftCenter
   
   
   Grid.TextMatrix(0, C_CAMPO) = "Campo de Información"
   Grid.TextMatrix(0, C_FORMATO) = "Formato"
   
End Sub

Private Sub LoadGrid()
   Dim i As Integer
   Dim j As Integer

   Grid.rows = Grid.FixedRows
   i = Grid.rows - 1
   
   For j = 0 To UBound(lFmtArray)
      Grid.rows = Grid.rows + 1
      i = i + 1
      Grid.TextMatrix(i, C_CAMPO) = lFmtArray(j).Campo
      Grid.TextMatrix(i, C_FORMATO) = lFmtArray(j).Formato
   Next j
   
   Call FGrVRows(Grid)
End Sub

Friend Sub FView(ByVal FmtCaption As String, FmtArray() As FmtImp_t)

   lFmtCaption = FmtCaption
   lFmtArray = FmtArray
   Me.Show vbModal

End Sub
Public Sub FViewEntidad()

   Call FillEntidad
   Me.Show vbModal

End Sub
Public Sub FViewLibCompras()

   Call FillLibroCompras
   Me.Show vbModal

End Sub
Public Sub FViewLibVentas()

   Call FillLibroVentas
   Me.Show vbModal

End Sub
Public Sub FViewLibReten()

   Call FillLibroReten
   Me.Show vbModal

End Sub
Public Sub FViewOtrosDocs()

   Call FillOtrosDocs
   Me.Show vbModal

End Sub

Public Sub FViewOtrosDocFull()

   Call FillOtrosDocFull
   Me.Show vbModal

End Sub

Public Sub FViewConfigCtasLibCompras()

   Call FillConfigCtasLibCompras
   Me.Show vbModal

End Sub

Public Sub FViewConfigCtasLibVentas()

   Call FillConfigCtasLibVentas
   Me.Show vbModal

End Sub

Private Sub FillEntidad()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Entidades"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "RUT o Referencia *"
   lFmtArray(i).Formato = "Si es RUT: con o sin punto y dígito verificador"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "No es RUT"
   lFmtArray(i).Formato = "Valor 1 indica que campo anterior NO ES RUT sino Referencia. Cero o blanco indica ES RUT."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Codigo *"
   lFmtArray(i).Formato = "Nombre corto de la entidad, en mayúscula y sin blancos, largo 15"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre o Razón Social *"
   lFmtArray(i).Formato = "Texto largo 80"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Dirección"
   lFmtArray(i).Formato = "Texto largo 100"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Comuna"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ciudad"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Teléfonos"
   lFmtArray(i).Formato = "Texto largo 30"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fax"
   lFmtArray(i).Formato = "Texto largo 15"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Giro"
   lFmtArray(i).Formato = "Texto largo 50"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Domicilio Postal"
   lFmtArray(i).Formato = "Texto largo 35"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Comuna Postal"
   lFmtArray(i).Formato = "Texto largo 20"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Email"
   lFmtArray(i).Formato = "Texto largo 50"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Sitio Web"
   lFmtArray(i).Formato = "Texto largo 50 "
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Observaciones"
   lFmtArray(i).Formato = "Texto largo 255"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Cliente"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Cliente"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Proveedor"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Proveedor"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Empleado"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Empleado"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Socio"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Socio"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Distribuidor"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es un Distribuidor"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Otro"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es Otros"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es Supermercado"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad es Supermercado"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   If gEmpresa.Ano >= 2020 Then
      lFmtArray(i).Campo = "Normas de Relación Art. 14D LIR"
   Else
      lFmtArray(i).Campo = "Normas de Relación Art. 14 TER "
   End If
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad se acoge estas normas."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Franquicia Tributaria"
   lFmtArray(i).Formato = "Número que indica franquicia tributaria (ver valores en Editar Entidad). Opcional en blanco."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Aplica Ret. 3% Prést. Solidario"
   lFmtArray(i).Formato = "Texto largo 1, valor 1 para indicar que la entidad aplica a retención"


End Sub
Public Sub FViewCuentas()

   lLbIFRS = True
   Call FillCuentas
   Me.Show vbModal

End Sub
Private Sub FillCuentas()
   Dim i As Integer

   lFmtCaption = "Formato Importación/Exportación de Plan de Cuentas"
   lIFRS = True
   
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Código"
   lFmtArray(i).Formato = "El formato depende de los niveles definidos. Ej.: 1-01-01-12"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre Corto"
   lFmtArray(i).Formato = "Texto largo 10, opcional en blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción *"
   lFmtArray(i).Formato = "Texto largo 100"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Doc. (RUT) asociado"
   lFmtArray(i).Formato = "Distinto de blanco, si debe llevar Documento o RUT asociado."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código FECU"
   lFmtArray(i).Formato = "Ej.: 5110000"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Capital Propio"
   lFmtArray(i).Formato = """Normal"", ""INTO"", ""CompActivo"", ""Exigible"", ""NoExigible"" (mayúsculas o minúsculas)"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Debe"
   lFmtArray(i).Formato = "Monto Debe para apertura cuenta"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Haber"
   lFmtArray(i).Formato = "Monto Haber para apertura cuenta"
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código F22"
   lFmtArray(i).Formato = "Código Formulario 22"
      
End Sub

Private Sub FillLibroCompras()
   Dim i As Integer

   lFmtCaption = "Formato de Captura de Documentos del Libro de Compras"
   lRepCaptura = True
   lNombCuentas = True
   lNulos = True
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Ingreso *"
   lFmtArray(i).Formato = "Fecha de ingreso del documento al libro. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento *"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (FAC, FCE, NCC, NDC, etc.)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, número <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. *"
   lFmtArray(i).Formato = "Número de documento"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Prop. IVA"
   lFmtArray(i).Formato = "Aplicar proporcionalidad de IVA (T (total), N (Nulo), P (Proporcional) o blanco)"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Emisión *"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT"
   lFmtArray(i).Formato = "RUT del emisior, con puntos (opcionalmente) y dígito verificador. Blanco para docs. sin entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Razón social"
   lFmtArray(i).Formato = "Texto largo 80, opcional en blanco para documentos que no requieren entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción *"
   lFmtArray(i).Formato = "Texto largo 100"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Sucursal"
   lFmtArray(i).Formato = "Texto largo 15. Si no se indica sucursal, el campo debe venir en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Afecto *"
   lFmtArray(i).Formato = "Valor Afecto, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Afecto"
   lFmtArray(i).Formato = "Código cuenta afecto con o sin guiones. Opcional en blanco."
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Exento *"
   lFmtArray(i).Formato = "Valor Exento, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Exento"
   lFmtArray(i).Formato = "Código cuenta exento con o sin guiones. Opcional en blanco."
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "IVA *"
   lFmtArray(i).Formato = "Valor IVA, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Otros Impuestos *"
   lFmtArray(i).Formato = "Valor Otros Impuestos, sin puntos ni comas. Positivo Imp. Adic. o Anticipos, Negativo Imp. Reten., cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Otros Impuestos"
   lFmtArray(i).Formato = "Código cuenta Otros Impuestos con o sin guiones. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Total *"
   lFmtArray(i).Formato = "Valor Total, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Total"
   lFmtArray(i).Formato = "Código cuenta total con o sin guiones. Opcional en blanco."
     
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Vencimiento *"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Interno"
   lFmtArray(i).Formato = "Numeración interna de la empresa para el documento. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Area Negocio"
   lFmtArray(i).Formato = "Código del área de negocio asociada al documento. Largo máx. 15. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión"
   lFmtArray(i).Formato = "Código del centro de gectión asociado al documento. Largo máx. 15. Opcional en blanco."
   

End Sub


Private Sub FillLibroVentas()
   Dim i As Integer

   lFmtCaption = "Formato de Captura de Documentos del Libro de Ventas"
   lRepCaptura = True
   lNombCuentas = True
   lNulos = True
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Emisión *"
   lFmtArray(i).Formato = "Fecha de emisión del documento. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento *"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (FAV, FVE, NCV, NDV, etc.)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Del Giro"
   lFmtArray(i).Formato = "Indica si es el documento es del giro o no. Blanco: ES del giro, <> blanco: NO es del giro"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, nro. <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Fiscal Impresora"
   lFmtArray(i).Formato = "N° Fiscal Impresora, sin puntos ni comas. Sólo para Máq. Registradora. Si no, en blanco."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Informe Z"
   lFmtArray(i).Formato = "N° Informe Z, sin puntos ni comas. Sólo para Máq. Registradora. Si no, en blanco."


   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. *"
   lFmtArray(i).Formato = "Número de documento"
   i = i + 1
   
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. Hasta"
   lFmtArray(i).Formato = "Número de documento hasta el cual incluye este registro (opcional en blanco). Utilizado para boletas"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cant. Boletas"
   lFmtArray(i).Formato = "Cantidad de boletas, sin puntos ni comas. Sólo para Vales de Pago Elect. Si no, en blanco."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT"
   lFmtArray(i).Formato = "RUT del receptor, con puntos (opcionalmente) y dígito verificador. Blanco para docs. sin entidad"

    i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Razón social"
   lFmtArray(i).Formato = "Texto largo 80, opcional en blanco para documentos que no requieren entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción *"
   lFmtArray(i).Formato = "Texto largo 100"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Sucursal"
   lFmtArray(i).Formato = "Texto largo 15. Si no se indica sucursal, el campo debe venir en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Afecto *"
   lFmtArray(i).Formato = "Valor Afecto, sin puntos ni comas. Siempre positivo o cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Afecto"
   lFmtArray(i).Formato = "Código cuenta afecto con o sin guiones. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Exento *"
   lFmtArray(i).Formato = "Valor Exento, sin puntos ni comas. Siempre positivo o cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Exento"
   lFmtArray(i).Formato = "Código cuenta exento con o sin guiones. Opcional en blanco."
   
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "IVA *"
   lFmtArray(i).Formato = "Valor IVA, sin puntos ni comas. Siempre positivo o cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Otros Impuestos *"
   lFmtArray(i).Formato = "Valor Otros Impuestos, sin puntos ni comas. Positivo Imp. Adic. o Anticipos, Negativo Imp. Reten., cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Otros Impuestos"
   lFmtArray(i).Formato = "Código cuenta Otros Impuestos con o sin guiones. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Total *"
   lFmtArray(i).Formato = "Valor Total, sin puntos ni comas. Siempre positivo o cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta Total"
   lFmtArray(i).Formato = "Código cuenta total con o sin guiones. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ventas Acum. Informe Z"
   lFmtArray(i).Formato = "Ventas acumuladas Informe Z, sin puntos ni comas. Sólo Tipo Doc. MRG. Si no, en blanco."
   
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Vencimiento. *"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Interno"
   lFmtArray(i).Formato = "Numeración interna de la empresa para el documento. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cód. Area Negocio "
   lFmtArray(i).Formato = "Código del área de negocio asociada al documento. Largo máx. 15. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cód. Centro Gestión "
   lFmtArray(i).Formato = "Código del centro de gectión asociado al documento. Largo máx. 15. Opcional en blanco."
   
   

End Sub

Private Sub FillLibroReten()
   Dim i As Integer
   Dim ImptoHon As Double

   lFmtCaption = "Formato de Captura de Documentos del Libro de Retenciones"
   lRepCaptura = True
   lNombCuentas = True
   lNulos = True
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Ingreso *"
   lFmtArray(i).Formato = "Fecha de ingreso del documento al libro. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento *"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (BOH, BRT)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, nro. <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. *"
   lFmtArray(i).Formato = "Número de documento"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Emisión *"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT *"
   lFmtArray(i).Formato = "RUT del emisor, con puntos (opcionalmente) y dígito verificador."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre *"
   lFmtArray(i).Formato = "Texto largo 80."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción *"
   lFmtArray(i).Formato = "Texto largo 100"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Sucursal"
   lFmtArray(i).Formato = "Texto largo 15. Si no se indica sucursal, el campo debe venir en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Honorarios *"
   lFmtArray(i).Formato = "Valor Honorarios sin Retención, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Bruto *"
   lFmtArray(i).Formato = "Valor Bruto, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "% Imp *"
   ImptoHon = ImpBolHono(DateSerial(gEmpresa.Ano, 1, 1)) * 100
   lFmtArray(i).Formato = "Porcentaje de impuesto, sin % (10, 20 u Otro). Para decimales usar " & W.CurDecSym & " (ej. " & Format(ImptoHon, DBLFMT2) & ")"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Impuesto *"
   lFmtArray(i).Formato = "Valor Impuesto, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Retención 3%"
   lFmtArray(i).Formato = "Valor Retención 3%, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Neto *"
   lFmtArray(i).Formato = "Valor Neto, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Tipo Retención *"
   lFmtArray(i).Formato = "Tipo de retención (Honorarios, Dieta u Otro)."
     
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta"
   lFmtArray(i).Formato = "Código cuenta Honorarios o Bruto. Opcional en blanco."

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Vencim."
   lFmtArray(i).Formato = "Fecha de vencimeinto del documento. Formato: dd/mm/aaaa"

End Sub


Private Sub FillOtrosDocs()
   Dim i As Integer

   lFmtCaption = "Formato de Captura de Otros Documentos"
   lRepCaptura = True
   lNombCuentas = True
   lNulos = False
   lOtrosDocs = True
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Emisión *"
   lFmtArray(i).Formato = "Fecha de emisión del documento. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento *"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (CHE, CHF, ABO, CAR, etc.)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, número <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. *"
   lFmtArray(i).Formato = "Número de documento"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT"
   lFmtArray(i).Formato = "RUT del emisior, con puntos (opcionalmente) y dígito verificador. Blanco para docs. sin entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Razón social"
   lFmtArray(i).Formato = "Texto largo 80, opcional en blanco para documentos que no requieren entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Observaciones"
   lFmtArray(i).Formato = "Texto largo 100, opcional en blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Valor *"
   lFmtArray(i).Formato = "Valor Total, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta"
   lFmtArray(i).Formato = "Código cuenta con o sin guiones. Opcional en blanco."
     
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Vencimiento"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Interno"
   lFmtArray(i).Formato = "Numeración interna de la empresa para el documento. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Incluir en Info. Analítico"
   lFmtArray(i).Formato = "Incluir el docuemtno en el informe Analítico. 0 o blanco: NO, 1: SI "
   
End Sub

Private Sub FillOtrosDocFull()
   Dim i As Integer

   lFmtCaption = "Formato de Captura de Otros Documentos Full"
   lRepCaptura = True
   lNombCuentas = True
   lNulos = False
   lOtrosDocs = True
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Emisión *"
   lFmtArray(i).Formato = "Fecha de emisión del documento. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento *"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (ODF)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, número <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc. *"
   lFmtArray(i).Formato = "Número de documento"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT"
   lFmtArray(i).Formato = "RUT del emisior, con puntos (opcionalmente) y dígito verificador. Blanco para docs. sin entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Razón social"
   lFmtArray(i).Formato = "Texto largo 80, opcional en blanco para documentos que no requieren entidad"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Observaciones"
   lFmtArray(i).Formato = "Texto largo 100, opcional en blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Valor *"
   lFmtArray(i).Formato = "Valor Total, sin puntos ni comas. Siempre positivo o en cero."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cod. Cuenta"
   lFmtArray(i).Formato = "Código cuenta con o sin guiones. Opcional en blanco."
     
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Vencimiento"
   lFmtArray(i).Formato = "Formato dd/mm/aaaa. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Interno"
   lFmtArray(i).Formato = "Numeración interna de la empresa para el documento. Opcional en blanco."
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Incluir en Info. Analítico"
   lFmtArray(i).Formato = "Incluir el documento en el informe Analítico. 0 o blanco: NO, 1: SI "
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Tratamiento *"
   lFmtArray(i).Formato = "Incluir el documento el Tratamiento. 1 o blanco: Activo, 2: Pasivo "
   
End Sub
Private Sub FillActivoFijo()
   Dim i As Integer

   lFmtCaption = "Formato de Captura de Activos Fijos"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "AF Totalmente Depreciado"
   lFmtArray(i).Formato = "Indica Activo Fijo totalmente depreciado: S/N/blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "AF no depreciable"
   lFmtArray(i).Formato = "Indica Activo Fijo no depreciable. S/N/blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Compra *"
   lFmtArray(i).Formato = "Fecha de compra. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Utilización"
   lFmtArray(i).Formato = "Fecha de utilización. Obligatorio si Act. Fijo es depreciable. Formato: dd/mm/aaaa"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cantidad *"
   lFmtArray(i).Formato = "Cantidad de unidades. Mayor que cero"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción *"
   lFmtArray(i).Formato = "Descripción del activo fijo"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Neto *"
   lFmtArray(i).Formato = "Neto compra"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "IVA *"
   lFmtArray(i).Formato = "IVA compra"
   i = i + 1
   
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Cred. 33 bis"
   lFmtArray(i).Formato = "Indica si se acoge a Crédito 33 bis: S/N/blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Valor Credito"
   lFmtArray(i).Formato = "Valor crédito 33 bis. Valor mayor que cero / blanco"

    i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Vida Útil"
   lFmtArray(i).Formato = "Vida útil del bien en meses. Valor mayor que cero"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Normal"
   lFmtArray(i).Formato = "Meses depreciación Normal, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Acelerada"
   lFmtArray(i).Formato = "Meses depreciación Acelerada, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Instantánea"
   lFmtArray(i).Formato = "Meses depreciación Instantánea, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Décima Parte"
   lFmtArray(i).Formato = "Meses depreciación Décima Parte, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Décima Parte MT"
   lFmtArray(i).Formato = "Meses depreciación Décima Parte MT, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ley 21.210 - Dep. Inst. e Inmed."
   lFmtArray(i).Formato = "Se acoge a Ley 21.210 Depreciación Instantánea e Inmediata, si aplica. S/N/blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ley 21.210 - Araucanía"
   lFmtArray(i).Formato = "Se acoge a Ley 21.210 Araucanía, si aplica. S/N/blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Ley 21.256 - Dep. Inst e Inmed."
   lFmtArray(i).Formato = "Se acoge a Ley 21.256 Instantanea e Inmediata, si aplica. S/N/blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Normal Hist."
   lFmtArray(i).Formato = "Meses depreciación normal histórica, si aplica. Valor > que cero o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Acel. Hist."
   lFmtArray(i).Formato = "Meses depreciación acelerada histórica, si aplica. Valor > que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Instant. Hist."
   lFmtArray(i).Formato = "Meses depreciación instantánea histórica, si aplica. Valor > que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Meses Dep. Décima Parte Hist."
   lFmtArray(i).Formato = "Meses depreciación décima parte histórica, si aplica. Valor > que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Valor Dep. Hist."
   lFmtArray(i).Formato = "Valor dep. histórica acumulada, si aplica. Valor > que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Venta o Baja"
   lFmtArray(i).Formato = "Indica si se realizó venta o baja del bien: V(venta)/B(baja)/blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Venta o Baja"
   lFmtArray(i).Formato = "Fecha de venta o baja del bien.  Formato: dd/mm/aaaa o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Neto Venta"
   lFmtArray(i).Formato = "Valor neto de venta. Valor mayor o igual que cero, o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "IVA Venta"
   lFmtArray(i).Formato = "Valor IVA venta. Valor mayor o igual que cero, o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta"
   lFmtArray(i).Formato = "Código cuenta contable, con o sin guiones. Opcional en blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Patente o Rol"
   lFmtArray(i).Formato = "Patente, Rol o Inscripción según proceda. Opcional. Texto largo máximo 30"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre Proyecto"
   lFmtArray(i).Formato = "Nombre Proyecto. Texto largo máximo 60. Opcional en blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Proyecto"
   lFmtArray(i).Formato = "Fecha Proyecto. Formato: dd/mm/aaaa o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Grupo"
   lFmtArray(i).Formato = "Nombre Grupo. Opcional en blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Incorporación"
   lFmtArray(i).Formato = "Fecha incorporación del bien.  Formato: dd/mm/aaaa o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Disponible"
   lFmtArray(i).Formato = "Fecha de disponibilidad del bien.  Formato: dd/mm/aaaa o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Derechos Internación"
   lFmtArray(i).Formato = "Valor Derechos de Internación. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Transporte"
   lFmtArray(i).Formato = "Valor Derechos de Transporte. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Adaptación"
   lFmtArray(i).Formato = "Valor Adaptación. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Otros Adquisición"
   lFmtArray(i).Formato = "Otros Gastos de Adquisición. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "IVA Recuperable"
   lFmtArray(i).Formato = "Valor IVA Recuperable. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Formación Personal"
   lFmtArray(i).Formato = "Valor de Formación del Personal. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Reubicación"
   lFmtArray(i).Formato = "Valor Reubicación. Valor mayor o igual que cero o blanco"
  
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Otros Gastos no Recon."
   lFmtArray(i).Formato = "Valor de otros gastos no reconocidos. Valor mayor o igual que cero o blanco"
   
   '2861733 tema 2
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Area Negocio."
   lFmtArray(i).Formato = "Código Area de Negocio. Opcional en blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestion."
   lFmtArray(i).Formato = "Código Centro de Gestion. Opcional en blanco"
   '2861733 tema 2

End Sub

Private Sub FillConfigCtasLibCompras()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Configuración Cuentas Libro de Compras"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "RUT Proveedor *"
   lFmtArray(i).Formato = "RUT: con o sin punto y dígito verificador"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre o Razón Social"
   lFmtArray(i).Formato = "Texto largo 80. Opcional si la entidad ya existe"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Prop. IVA"
   lFmtArray(i).Formato = "Proporcionalidad de IVA. Valores válidos: blanco, T, N, P"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Afecto"
   lFmtArray(i).Formato = "Código cuenta afecto, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Exento"
   lFmtArray(i).Formato = "Código cuenta exento, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Total"
   lFmtArray(i).Formato = "Código cuenta total, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Afecto"
   lFmtArray(i).Formato = "Código área de negocio afecto, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Exento"
   lFmtArray(i).Formato = "Código área de negocio exento, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Total"
   lFmtArray(i).Formato = "Código área de negocio total, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Afecto"
   lFmtArray(i).Formato = "Código centro de gestión afecto, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Exento"
   lFmtArray(i).Formato = "Código centro de gestión exento, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Total"
   lFmtArray(i).Formato = "Código centro de gestión total, o blanco"



End Sub

Private Sub FillConfigCtasLibVentas()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Configuración Cuentas Libro de Ventas"
   
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "RUT Cliente *"
   lFmtArray(i).Formato = "RUT: con o sin punto y dígito verificador"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Nombre o Razón Social"
   lFmtArray(i).Formato = "Texto largo 80. Opcional si la entidad ya existe"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Es del Giro"
   lFmtArray(i).Formato = "Indicador si venta es del Giro. Valores válidos: S, N"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Afecto"
   lFmtArray(i).Formato = "Código cuenta afecto, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Exento"
   lFmtArray(i).Formato = "Código cuenta exento, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Cuenta Total"
   lFmtArray(i).Formato = "Código cuenta total, con o sin guión, o blanco"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Afecto"
   lFmtArray(i).Formato = "Código área de negocio afecto, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Exento"
   lFmtArray(i).Formato = "Código área de negocio exento, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área Negocio Total"
   lFmtArray(i).Formato = "Código área de negocio total, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Afecto"
   lFmtArray(i).Formato = "Código centro de gestión afecto, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Exento"
   lFmtArray(i).Formato = "Código centro de gestión exento, o blanco"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro Gestión Total"
   lFmtArray(i).Formato = "Código centro de gestión total, o blanco"



End Sub

Public Sub FViewComprobantes()

   lEjemplo = True
   Call FillComprobantes
   Me.Show vbModal

End Sub
Public Sub FViewActivoFijo()

   lEjemplo = False
   Call FillActivoFijo
   Me.Show vbModal

End Sub
Private Sub FillComprobantes()
   Dim i As Integer

   lFmtCaption = "Formato Importación de Comprobantes"
   lNombCuentas = True
        
   i = 0
   ReDim lFmtArray(i)
   lFmtArray(i).Campo = "Tipo Comprobante *"
   lFmtArray(i).Formato = """Ingreso"", ""Egreso"" o ""Traspaso"" para primer registro y blanco para registros siguientes"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Fecha Ingreso *"
   lFmtArray(i).Formato = "dd/mm/aaaa para primer registro y blanco para registros siguientes"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Total Comprobante *"
   lFmtArray(i).Formato = "Valor mayor que cero, sin puntos ni comas, para primer registro y blanco para registros siguientes"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Estado comprobante *"
   lFmtArray(i).Formato = """Pendiente"", ""Aprobado"" o ""Anulado"" para primer registro y blanco para registros siguientes"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Glosa comprobante *"
   lFmtArray(i).Formato = "Texto distinto de blanco para primer registro y blanco para registros siguientes"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código cuenta *"
   lFmtArray(i).Formato = "Código cuenta movimiento, con o sin guiones"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Debe *"
   lFmtArray(i).Formato = "Valor Debe, sin puntos ni comas. Siempre positivo o en cero"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Haber *"
   lFmtArray(i).Formato = "Valor Haber, sin puntos ni comas. Siempre positivo o en cero"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Descripción movimiento"
   lFmtArray(i).Formato = "Texto opcional asociado al movimiento"
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Área de Negocio"
   lFmtArray(i).Formato = "Texto opcional. Debe ser exactamente igual al código ingresado en el sistema"
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Código Centro de Gestión"
   lFmtArray(i).Formato = "Texto opcional. Debe ser exactamente igual al código ingresado en el sistema"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "Tipo de Libro"
   lFmtArray(i).Formato = "Inicial de tipo de libro (C:Compras, V:Ventas, R:Retenciones, O:Otros Docs., S:Sueldos/Remun., F:Otros Docs Full.)"
   
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "(TD) Tipo de Documento"
   lFmtArray(i).Formato = "Diminutivo de tipo de documento en mayúsculas (FAC, FAV, FCE, FVE, NCC, NCV, NDC, ODF, etc.)"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "DTE"
   lFmtArray(i).Formato = "Indica si es DTE (Doc. Tributario Electrónico). 0 o blanco: NO es DTE, número <> 0: SI es DTE"

   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "N° Doc."
   lFmtArray(i).Formato = "Número de documento"
      
   i = i + 1
   ReDim Preserve lFmtArray(i)
   lFmtArray(i).Campo = "RUT"
   lFmtArray(i).Formato = "RUT del emisior, con puntos (opcionalmente) y dígito verificador. Blanco para docs. sin entidad"

      
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption)
   End If
      
End Sub
