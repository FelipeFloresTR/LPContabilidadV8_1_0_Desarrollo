VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Empresa"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
   Icon            =   "FrmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab Tab1 
      Height          =   6915
      Left            =   1140
      TabIndex        =   74
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12197
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Antecedentes Empresa"
      TabPicture(0)   =   "FrmEmpresa.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Bt_TareDatosAnoAnt"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Antecedentes  Legales"
      TabPicture(1)   =   "FrmEmpresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Im_Exc(0)"
      Tab(1).Control(2)=   "Label1(8)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tipo de Contribuyente"
      TabPicture(2)   =   "FrmEmpresa.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Fr_TrBolsa"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Fr_LibCaja"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Propietarios y Socios"
      TabPicture(3)   =   "FrmEmpresa.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Bt_Calc"
      Tab(3).Control(1)=   "Bt_Del"
      Tab(3).Control(2)=   "Bt_Print"
      Tab(3).Control(3)=   "Bt_CopyExcel"
      Tab(3).Control(4)=   "Tx_CurrCell"
      Tab(3).Control(5)=   "Grid"
      Tab(3).Control(6)=   "GridTot"
      Tab(3).Control(7)=   "Label3"
      Tab(3).Control(8)=   "Lb_Socios"
      Tab(3).ControlCount=   9
      Begin VB.CommandButton Bt_TareDatosAnoAnt 
         Caption         =   "Traer Datos Año Anterior"
         Height          =   1095
         Left            =   -65100
         Picture         =   "FrmEmpresa.frx":007C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   660
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Calc 
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
         Left            =   -73560
         Picture         =   "FrmEmpresa.frx":06EF
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Calculadora"
         Top             =   420
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
         Left            =   -74880
         Picture         =   "FrmEmpresa.frx":0A50
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Eliminar movimiento seleccionado"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Bt_Print 
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
         Left            =   -74400
         Picture         =   "FrmEmpresa.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Imprimir"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
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
         Left            =   -73980
         Picture         =   "FrmEmpresa.frx":1306
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Copiar Excel"
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox Tx_CurrCell 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   6060
         Width           =   6255
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   67
         Top             =   840
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8599
         Cols            =   2
         Rows            =   4
         FixedCols       =   1
         FixedRows       =   2
         ScrollBars      =   3
         AllowUserResizing=   0
         HighLight       =   1
         SelectionMode   =   0
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   -1  'True
         Locked          =   0   'False
      End
      Begin VB.Frame Fr_LibCaja 
         Caption         =   "Libro de Ingreso-Egreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   8040
         TabIndex        =   79
         Top             =   720
         Width           =   3075
         Begin VB.TextBox Ch_ObligaLibComprasVentas2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   118
            Text            =   "Ley de IVA"
            Top             =   840
            Width           =   2325
         End
         Begin VB.CheckBox Ch_ObligaLibComprasVentas 
            Caption         =   "Se encuentra obligado a llevar"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   2715
         End
         Begin VB.TextBox Ch_ObligaLibComprasVentas1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   117
            Text            =   "Libro Compras Ventas según la"
            Top             =   600
            Width           =   2500
         End
      End
      Begin VB.Frame Fr_TrBolsa 
         Caption         =   "Transa en la Bolsa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   540
         TabIndex        =   78
         Top             =   5760
         Width           =   2895
         Begin VB.OptionButton Op_TrBolsaNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1800
            TabIndex        =   49
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Op_TrBolsaSi 
            Caption         =   "Si"
            Height          =   255
            Left            =   540
            TabIndex        =   48
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Franquicias Tributarias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   3900
         TabIndex        =   77
         Top             =   660
         Width           =   3675
         Begin VB.Frame Fr_FranqNueva2020 
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   3435
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "No sujeto art. 14 LIR"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   65
               Top             =   2280
               Width           =   3120
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Otro"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   64
               Top             =   1920
               Width           =   3120
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "14 B N° 1 Renta efectiva sin Balance"
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   63
               Top             =   1560
               Width           =   3120
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Rentas Presuntas"
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   62
               Top             =   1200
               Width           =   3120
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "14 D N°8 Régimen Pro Pyme Transp."
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   61
               Top             =   840
               Width           =   3000
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "14 D N°3 Régimen Pro Pyme General"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   3120
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "14 A Régimen Semi Integrado"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   59
               Top             =   120
               Width           =   3000
            End
         End
         Begin VB.Frame Fr_FranqComun 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   60
            TabIndex        =   122
            Top             =   2820
            Width           =   3195
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Soc. Prof. 1ra. Categoría"
               Height          =   255
               Index           =   10
               Left            =   180
               TabIndex        =   124
               Top             =   60
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Soc. Prof. 2da. Categoría"
               Height          =   255
               Index           =   11
               Left            =   180
               TabIndex        =   123
               Top             =   420
               Width           =   2400
            End
         End
         Begin VB.Frame Fr_FranqOriginal 
            BorderStyle     =   0  'None
            Height          =   3315
            Left            =   120
            TabIndex        =   121
            Top             =   180
            Width           =   3435
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Régimen Artículo 14 bis"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   50
               Top             =   180
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Ley 18.392 / 19.149"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   55
               Top             =   1980
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "D. L. 600"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   56
               Top             =   2340
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "D. L. 701"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   57
               Top             =   2700
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "D. S. 341"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   58
               Top             =   3060
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Régimen Artículo 14 Ter A)"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   53
               Top             =   1260
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Régimen Artículo 14 quater"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   54
               Top             =   1620
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Régimen Renta Atribuida"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   51
               Top             =   540
               Width           =   2400
            End
            Begin VB.CheckBox Ch_Franquicia 
               Caption         =   "Régimen Semi Integrado"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   52
               Top             =   900
               Width           =   2400
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Contribuyente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   540
         TabIndex        =   76
         Top             =   720
         Width           =   2895
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Org. sin Fines de Lucro"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   47
            Top             =   3960
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Cooperativas"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   46
            Top             =   3600
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Comunidad"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   45
            Top             =   3240
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Establecimiento Permanente"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   44
            Top             =   2880
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad de Profesionales"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   43
            Top             =   2520
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Empresario Individual"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   42
            Top             =   2160
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Empresario Individual (EIRL)"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   41
            Top             =   1800
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Soc. Personas 1ª Categoría"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad por Acción"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad Anónima Cerrada"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   2475
         End
         Begin VB.OptionButton Op_TipoContrib 
            Caption         =   "Sociedad Anónima Abierta"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   2475
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Información básica de la empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -74580
         TabIndex        =   80
         Top             =   600
         Width           =   8955
         Begin VB.TextBox Txt_VillaOPob 
            Height          =   315
            Left            =   420
            MaxLength       =   50
            TabIndex        =   8
            Top             =   2805
            Width           =   8115
         End
         Begin VB.TextBox Txt_CodArea 
            Height          =   315
            Left            =   420
            MaxLength       =   4
            TabIndex        =   12
            Top             =   3915
            Width           =   975
         End
         Begin VB.TextBox Txt_Celular 
            Height          =   315
            Left            =   3720
            MaxLength       =   9
            TabIndex        =   14
            Top             =   3915
            Width           =   2175
         End
         Begin VB.CommandButton Bt_Email 
            Height          =   375
            Left            =   3960
            Picture         =   "FrmEmpresa.frx":174B
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   5055
            Width           =   375
         End
         Begin VB.TextBox Tx_EMail 
            Height          =   315
            Left            =   420
            MaxLength       =   50
            TabIndex        =   18
            Top             =   5055
            Width           =   3555
         End
         Begin VB.TextBox tx_NombreCorto 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   555
            Width           =   2475
         End
         Begin VB.TextBox Tx_Fax 
            Height          =   315
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   15
            Top             =   3915
            Width           =   2535
         End
         Begin VB.TextBox Tx_Telefonos 
            Height          =   315
            Left            =   1500
            MaxLength       =   11
            TabIndex        =   13
            Top             =   3915
            Width           =   2055
         End
         Begin VB.TextBox Tx_Ciudad 
            Height          =   315
            Left            =   6000
            MaxLength       =   20
            TabIndex        =   11
            Top             =   3375
            Width           =   2535
         End
         Begin VB.TextBox Tx_Calle 
            Height          =   315
            Left            =   420
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2235
            Width           =   5715
         End
         Begin VB.TextBox Tx_DirPostal 
            Height          =   315
            Left            =   420
            MaxLength       =   30
            TabIndex        =   16
            Top             =   4455
            Width           =   5535
         End
         Begin VB.TextBox Tx_RazonSocial 
            Height          =   315
            Left            =   420
            MaxLength       =   200
            TabIndex        =   2
            Top             =   1155
            Width           =   8115
         End
         Begin VB.TextBox Tx_RUT 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   420
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   555
            Width           =   1395
         End
         Begin VB.TextBox Tx_Numero 
            Height          =   315
            Left            =   6180
            MaxLength       =   15
            TabIndex        =   6
            Top             =   2220
            Width           =   1155
         End
         Begin VB.TextBox Tx_Dpto 
            Height          =   315
            Left            =   7380
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2220
            Width           =   1155
         End
         Begin VB.TextBox Tx_ApMaterno 
            Height          =   315
            Left            =   420
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1695
            Width           =   4035
         End
         Begin VB.TextBox Tx_Nombre 
            Height          =   315
            Left            =   4500
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1695
            Width           =   4035
         End
         Begin VB.ComboBox Cb_Region 
            Height          =   315
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3375
            Width           =   2775
         End
         Begin VB.ComboBox Cb_Comuna 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3375
            Width           =   2715
         End
         Begin VB.ComboBox Cb_ComPostal 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   4455
            Width           =   2535
         End
         Begin VB.TextBox Tx_Web 
            Height          =   315
            Left            =   4440
            MaxLength       =   50
            TabIndex        =   20
            Top             =   5055
            Width           =   3675
         End
         Begin VB.CommandButton Bt_Web 
            Height          =   375
            Left            =   8100
            Picture         =   "FrmEmpresa.frx":1B56
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   5055
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Villa o Población:"
            Height          =   195
            Index           =   30
            Left            =   420
            TabIndex        =   127
            Top             =   2595
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Área"
            Height          =   195
            Index           =   29
            Left            =   420
            TabIndex        =   126
            Top             =   3720
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Celular:"
            Height          =   195
            Index           =   28
            Left            =   3720
            TabIndex        =   125
            Top             =   3720
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre corto:"
            Height          =   195
            Index           =   13
            Left            =   2040
            TabIndex        =   98
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "E-Mail:"
            Height          =   195
            Index           =   16
            Left            =   420
            TabIndex        =   97
            Top             =   4860
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Index           =   10
            Left            =   6000
            TabIndex        =   96
            Top             =   3720
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Teléfonos:"
            Height          =   195
            Index           =   9
            Left            =   1500
            TabIndex        =   95
            Top             =   3720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Región:"
            Height          =   195
            Index           =   7
            Left            =   420
            TabIndex        =   94
            Top             =   3180
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio Postal:"
            Height          =   195
            Index           =   6
            Left            =   420
            TabIndex        =   93
            Top             =   4260
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad:"
            Height          =   195
            Index           =   5
            Left            =   6000
            TabIndex        =   92
            Top             =   3180
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comuna:"
            Height          =   195
            Index           =   4
            Left            =   3240
            TabIndex        =   91
            Top             =   3195
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Calle:"
            Height          =   195
            Index           =   3
            Left            =   420
            TabIndex        =   90
            Top             =   2040
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social/Apellido Paterno:"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   89
            Top             =   960
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RUT:"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   88
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Index           =   17
            Left            =   6180
            TabIndex        =   87
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Of. Dpto:"
            Height          =   195
            Index           =   18
            Left            =   7380
            TabIndex        =   86
            Top             =   2040
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Materno:"
            Height          =   195
            Index           =   19
            Left            =   420
            TabIndex        =   85
            Top             =   1500
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombres:"
            Height          =   195
            Index           =   20
            Left            =   4500
            TabIndex        =   84
            Top             =   1500
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comuna Postal:"
            Height          =   195
            Index           =   21
            Left            =   6000
            TabIndex        =   83
            Top             =   4260
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sitio Web:"
            Height          =   195
            Index           =   22
            Left            =   4440
            TabIndex        =   82
            Top             =   4860
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "(*) Comuna descontinuada por el SII"
            Height          =   195
            Left            =   420
            TabIndex        =   81
            Top             =   5475
            Width           =   2550
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5715
         Left            =   -74460
         TabIndex        =   99
         Top             =   660
         Width           =   8895
         Begin VB.TextBox Tx_CodActEcon 
            Height          =   315
            Left            =   7200
            MaxLength       =   6
            TabIndex        =   29
            Top             =   1755
            Width           =   1275
         End
         Begin VB.ComboBox Cb_ActEcon 
            Height          =   315
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1755
            Width           =   6735
         End
         Begin VB.Frame Frame2 
            Caption         =   "Datos Contador"
            ForeColor       =   &H00FF0000&
            Height          =   975
            Index           =   2
            Left            =   420
            TabIndex        =   105
            Top             =   2235
            Width           =   8055
            Begin VB.TextBox Tx_RutContador 
               Height          =   315
               Left            =   180
               MaxLength       =   12
               TabIndex        =   30
               Top             =   480
               Width           =   1155
            End
            Begin VB.TextBox Tx_Contador 
               Height          =   315
               Left            =   1440
               MaxLength       =   30
               TabIndex        =   31
               Top             =   480
               Width           =   6420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RUT:"
               Height          =   195
               Index           =   23
               Left            =   180
               TabIndex        =   107
               Top             =   290
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nombre:"
               Height          =   195
               Index           =   24
               Left            =   1440
               TabIndex        =   106
               Top             =   290
               Width           =   600
            End
         End
         Begin VB.TextBox Tx_Giro 
            Height          =   315
            Left            =   420
            MaxLength       =   80
            TabIndex        =   23
            Top             =   555
            Width           =   8040
         End
         Begin VB.TextBox Tx_FConstit 
            Height          =   315
            Left            =   2940
            TabIndex        =   24
            Top             =   1035
            Width           =   1035
         End
         Begin VB.TextBox Tx_FInicioAct 
            Height          =   315
            Left            =   7200
            TabIndex        =   26
            Top             =   1035
            Width           =   1035
         End
         Begin VB.Frame Frame1 
            Caption         =   "Representantes Legales"
            ForeColor       =   &H00FF0000&
            Height          =   1755
            Left            =   420
            TabIndex        =   100
            Top             =   3315
            Width           =   8055
            Begin VB.TextBox Tx_RUTRep 
               Height          =   315
               Index           =   0
               Left            =   180
               MaxLength       =   13
               TabIndex        =   33
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox Tx_NombreRep 
               Height          =   315
               Index           =   0
               Left            =   1440
               MaxLength       =   30
               TabIndex        =   34
               Top             =   600
               Width           =   6420
            End
            Begin VB.TextBox Tx_NombreRep 
               Height          =   315
               Index           =   1
               Left            =   1440
               MaxLength       =   30
               TabIndex        =   36
               Top             =   1200
               Width           =   6420
            End
            Begin VB.TextBox Tx_RUTRep 
               Height          =   315
               Index           =   1
               Left            =   180
               MaxLength       =   13
               TabIndex        =   35
               Top             =   1200
               Width           =   1155
            End
            Begin VB.CheckBox Ch_RepConjunta 
               Caption         =   "Representación conjunta"
               Height          =   255
               Left            =   5760
               TabIndex        =   32
               Top             =   300
               Width           =   2115
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RUT:"
               Height          =   195
               Index           =   11
               Left            =   180
               TabIndex        =   104
               Top             =   405
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nombre:"
               Height          =   195
               Index           =   12
               Left            =   1440
               TabIndex        =   103
               Top             =   405
               Width           =   600
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nombre:"
               Height          =   195
               Index           =   14
               Left            =   1440
               TabIndex        =   102
               Top             =   1005
               Width           =   600
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RUT:"
               Height          =   195
               Index           =   15
               Left            =   180
               TabIndex        =   101
               Top             =   1005
               Width           =   390
            End
         End
         Begin VB.CommandButton Bt_FConstit 
            Height          =   315
            Left            =   3960
            Picture         =   "FrmEmpresa.frx":1E60
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1035
            Width           =   225
         End
         Begin VB.CommandButton Bt_FInicioAct 
            Height          =   315
            Left            =   8220
            Picture         =   "FrmEmpresa.frx":216A
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1035
            Width           =   225
         End
         Begin VB.Image Im_Exc 
            Height          =   330
            Index           =   1
            Left            =   420
            Picture         =   "FrmEmpresa.frx":2474
            Top             =   5100
            Width           =   300
         End
         Begin VB.Label La_Url 
            Caption         =   "www.sii.cl/catastro/codigos_economica.htm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   5220
            MouseIcon       =   "FrmEmpresa.frx":283E
            MousePointer    =   99  'Custom
            TabIndex        =   113
            Top             =   5160
            Width           =   3315
         End
         Begin VB.Label lb_MsgCodAct 
            Caption         =   "Código Act. Economica descontinuado, verifique su código en"
            Height          =   255
            Left            =   720
            TabIndex        =   112
            Top             =   5175
            Width           =   4455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clasificador de Actividades Económicas:"
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   111
            Top             =   1560
            Width           =   2865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giro:"
            Height          =   195
            Index           =   25
            Left            =   420
            TabIndex        =   110
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio Actividades:"
            Height          =   195
            Index           =   26
            Left            =   5340
            TabIndex        =   109
            Top             =   1095
            Width           =   1785
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Constitución de la Empresa:"
            Height          =   195
            Index           =   27
            Left            =   420
            TabIndex        =   108
            Top             =   1095
            Width           =   2460
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridTot 
         Height          =   315
         Left            =   -74880
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   5700
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   556
         _Version        =   393216
         Cols            =   19
         FixedCols       =   2
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NOTA: Configuar cuentas contables distintas para cada socio/accionistas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   119
         Top             =   6480
         Width           =   5235
      End
      Begin VB.Label Lb_Socios 
         Alignment       =   2  'Center
         Caption         =   "Propietarios, Socios, Comuneros o Accionistas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74820
         TabIndex        =   114
         Top             =   540
         Width           =   11295
      End
      Begin VB.Image Im_Exc 
         Height          =   330
         Index           =   0
         Left            =   -66720
         Picture         =   "FrmEmpresa.frx":2990
         Top             =   2580
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód. act. ecón.:"
         Height          =   195
         Index           =   8
         Left            =   -68040
         TabIndex        =   75
         Top             =   2390
         Width           =   1140
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   13140
      TabIndex        =   68
      Top             =   420
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   13140
      TabIndex        =   69
      Top             =   780
      Width           =   1155
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   570
      Index           =   0
      Left            =   300
      Picture         =   "FrmEmpresa.frx":2D5A
      Top             =   480
      Width           =   570
   End
End
Attribute VB_Name = "FrmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDSOCIO = 0
Const C_RUT = 1
Const C_NOMBRE = 2
Const C_PJEPART = 3
Const C_CANTACCIONES = 4
Const C_MONTOSUSCRITO = 5
Const C_MONTOPAGADO = 6
Const C_IDCTAAPORTES = 7
Const C_CODCTAAPORTES = 8
Const C_CTAAPORTES = 9
Const C_IDCTARETIROS = 10
Const C_CODCTARETIROS = 11
Const C_CTARETIROS = 12
Const C_IDTIPOSOCIO = 13
Const C_TIPOSOCIO = 14
Const C_UPDATE = 15

Const NCOLS = C_UPDATE

Const VERSION_1 = 1

Const TAB_ANTEMP = 0
Const TAB_ANTLEG = 1
Const TAB_TCONTRIB = 2
Const TAB_SOCIOS = 3


Dim lRc As Integer
Dim lOper As Integer
Dim lcbCodActiv As ClsCombo
Dim lInLoad As Boolean
Dim lOrientacion As Integer
Dim lGetFromHR As Boolean
Dim lTipoContrib As Integer
Dim FranqActual As Integer
Dim x As Integer


'pipe lpremu

Const SG_PASSW_FAIRPAY = "oP,*/'#2j7h7_$3"

'Public lDbRemu As Database

Dim lEsLPRemu As Boolean
Dim lRemuSQLServer As Boolean

Dim lMsgAdv As Boolean
Dim lCtasRemu(MAX_CTASREMU) As Long
Dim lIdEmpresaRem As Long
Dim lPathlDbRemu As String
Dim lConnStr As String
Dim lEmpSep As Boolean
Dim lCbCCosto As ClsCombo
Dim lDesglozarCCosto As Boolean


'fin lpremu

Private Sub Bt_Calc_Click()
   Call Calculadora

End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_CopyExcel_Click()
   
   Call FGr2Clip(Grid, Lb_Socios)

End Sub

Private Sub Bt_Email_Click()
   Dim Buf As String
   Dim Rc As Long
   Dim Pos As Integer
   
   Pos = InStr(Tx_email, "@")
   If Trim(Tx_email) <> "" And Trim(Tx_RazonSocial) <> "" And Pos <> 0 Then
     Buf = "mailto:" & Trim(Tx_RazonSocial) & "<" & Trim(Tx_email) & ">"
     Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
     
   End If
   
End Sub

Private Sub Bt_FConstit_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FConstit)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_FInicioAct_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FInicioAct)
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()
   
   If valida() = True Then
      Call SaveAll
      Call ReadDatosBasEmpresa
      Unload Me
   End If
   
End Sub
Public Function FView(ByVal IdEmpresa As Long) As Integer
   lOper = O_VIEW
   Me.Show vbModal
   FView = lRc
End Function
Public Function FEdit(ByVal IdEmpresa As Long, Optional ByVal GetFromHR As Boolean = False) As Integer
   lOper = O_EDIT
   lGetFromHR = GetFromHR
   Me.Show vbModal
   FEdit = lRc
End Function

Private Sub Bt_Print_Click()
   
   If SelPrinter() Then
      Exit Sub
   End If
           
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = lOrientacion
   
   Call ResetPrtBas(gPrtReportes)
   MousePointer = vbDefault

End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   Dim Total(NCOLS) As String
   
   lOrientacion = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Lb_Socios
   gPrtReportes.Titulos = Titulos
    
   gPrtReportes.Encabezados = Encabezados
      
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
               
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_RUT
   
   gPrtReportes.NTotLines = 1
   gPrtReportes.Total = Total
   

End Sub

Private Sub Bt_TareDatosAnoAnt_Click()

   If MsgBox1("Desea traer los antecedentes de esta empresa, incluyendo Antecedentes Legales y Tipo de Contribuyente, almacenados en el año anterior?" & vbCrLf & vbCrLf & "ATENCIÓN: Esta información reemplazará la ya exiestente en este formulario.", vbQuestion + vbYesNoCancel) <> vbYes Then
      Exit Sub
   End If
   
#If DATACON = 1 Then
   If Not ExistFile(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb") Then
      MsgBox1 "No existe año anterior para esta empresa", vbExclamation
      Exit Sub
   End If
   
   'existe año anterior, linkeamos tabla Empresa
   
   'cerramos el año actual y abrimos el año anterior
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
   Call LinkMdbAdm
   
   'corrige base del año anterior, por si las moscas
   Call CorrigeBase

   'cerramos el año anterior  y abrimos el año actual
   Call CloseDb(DbMain)
   
   Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
   
   Call LinkMdbTable(DbMain, gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "Empresa", "EmpresaAnt", True, , gEmpresa.ConnStr)

   Call LoadAll("EmpresaAnt", -1, -1)   'en la tabla Empresa en Access hay un sólo registro, así que no se necesita ni IdEmpresa ni Ano
   
   Call ExecSQL(DbMain, "DROP TABLE EmpresaAnt")
   
#Else

   Call LoadAll("Empresa", gEmpresa.id, gEmpresa.Ano - 1)
   
#End If
   
End Sub


Private Sub Bt_Web_Click()
   Dim Rc As Long
   
   If Trim(Tx_Web) <> "" Then
      Rc = ShellExecute(Me.hWnd, "open", Tx_Web, "", "", 1)
   End If
   
End Sub

Private Sub Cb_ActEcon_Click()
   'Tx_CodActEcon = Right("000000" & ItemData(Cb_ActEcon), 5)
   
   'PS se cambio códgo de Actividad
'   lb_MsgCodAct = IIf(Val(lcbCodActiv.Matrix(2)) = VERSION_1 And gEmpresa.Ano >= 2005, "Código de Actividad Económica descontinuado, verifique su código en ", "")
'   Im_Exc(0).Visible = IIf(Val(lcbCodActiv.Matrix(2)) = VERSION_1 And gEmpresa.Ano >= 2005, True, False)
'   Im_Exc(1).Visible = Im_Exc(0).Visible
'   La_Url.Visible = Im_Exc(0).Visible
   
   Tx_CodActEcon = Right("000000" & lcbCodActiv.ItemData, 6)
   
End Sub

Private Sub Cb_Comuna_Click()
   Call SelItem(Cb_ComPostal, ItemData(Cb_Comuna))
End Sub

Private Sub Cb_Region_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Cod As String
   
   Cod = Right("00" & ItemData(Cb_Region), 2)
   
   Q1 = "SELECT Comuna, id FROM Regiones"
   Q1 = Q1 & " WHERE Codigo = '" & Cod & "'"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_Comuna.Clear
   Cb_Comuna.AddItem "<Ninguna>"
   Cb_Comuna.ItemData(Cb_Comuna.NewIndex) = 0
   Call FillCombo(Cb_Comuna, DbMain, Q1, -1, True)
   
End Sub

Private Sub Ch_Franquicia_Click(Index As Integer)
   
   If lInLoad Then
      Exit Sub
   End If
   
   If Index = FRANQ_SOCPROFPRIMCAT Then
      If Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) <> 0 Then
         Ch_Franquicia(FRANQ_SOCPROFSEGCAT) = 0
      End If
   ElseIf Index = FRANQ_SOCPROFSEGCAT Then
      If Ch_Franquicia(FRANQ_SOCPROFSEGCAT) <> 0 Then
         Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) = 0
      End If
   ElseIf Index = FRANQ_14TER Then
      Call MsgLey21210
   End If
      
'   Call EnabFranq

End Sub

Private Sub Ch_ObligaLibComprasVentas_Click()
   If lInLoad Then
      Exit Sub
   End If

   If Ch_ObligaLibComprasVentas <> 0 Then
      MsgBox1 "ATENCIÓN: Si marca esta opción NO podrá ingresar manualmente nuevos documentos al libro de caja, sólo podrá traerlos desde los libros de Compras y Ventas", vbInformation
   Else
      MsgBox1 "ATENCIÓN: Si desmarca esta opción, deberá ingresar manualmente los documentos al libro de caja y no podrá traerlos desde el Libro de Compras y Ventas", vbInformation
   End If

End Sub

Private Sub Form_Load()
   Tab1.Tab = 0
   Dim x As Integer
   lInLoad = True
     
   Call FillCombosFrm
   Call SetUpGrid
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   If gEmpresa.Ano < 2020 Then
      Fr_FranqOriginal.visible = True
      Fr_FranqNueva2020.visible = False
      Fr_FranqComun.Top = 3480
   Else
      Fr_FranqOriginal.visible = False
      Fr_FranqNueva2020.visible = True
      Fr_FranqComun.Top = 2820
   End If
   
   Call LoadAll
   
   Call SetTxRO(Tx_Rut, True)
   Call SetTxRO(Tx_NombreCorto, True)
   Call SetTxRO(Tx_CurrCell, True)

   If gAppCode.Demo Then
      Call SetTxRO(Tx_RazonSocial, True)
      Call SetTxRO(Tx_Nombre, True)
      Call SetTxRO(Tx_ApMaterno, True)
   End If
   
   Call SetTxRO(Ch_ObligaLibComprasVentas1, True)
   Call SetTxRO(Ch_ObligaLibComprasVentas2, True)
   
'   Call EnabFranq
      
   Call SetupPriv
   
   lInLoad = False
   
   For x = 1 To MAX_FRANQ
    If CBool(Ch_Franquicia(x)) Then
        FranqActual = x
    End If
   Next x
   
    Call LoadAll
     
End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Cod As String
   Dim UltimoNivel As Boolean
   Dim NombCta As String, DescCta As String
   Dim IdCuenta As Long
   Dim PjePart As Single

   Action = vbOK
   Value = Trim(Value)
   
   Select Case Col
   
      Case C_RUT
                  
         If Value = "" Then
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbOK
         
         ElseIf Value = "0-0" Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Action = vbRetry
         
         ElseIf Not ValidCID(Value) Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Action = vbRetry
         
         Else
            Value = FmtCID(vFmtCID(Value))
            Grid.TextMatrix(Row, C_RUT) = Value
            
         End If
         
      Case C_PJEPART
      
         If Value <> "" Then
            
            PjePart = vFmt(Value)
         
            If PjePart > 0 And PjePart <= 100 Then
               Value = Format(Value, DBLFMT2)
               Grid.TextMatrix(Row, Col) = Value
               Call CalcTotSocios
            Else
               Action = vbCancel
               MsgBox1 "Valor inválido.", vbExclamation
               Exit Sub
            End If
         End If
         
      Case C_CANTACCIONES
         If Value <> "" Then
            Value = Format(vFmt(Value), NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            Call CalcTotSocios
         End If
         
      Case C_MONTOSUSCRITO, C_MONTOPAGADO
         If Value <> "" Then
            Value = Format(vFmt(Value), NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            Call CalcTotSocios
         End If
         
      Case C_CODCTAAPORTES, C_CODCTARETIROS
         Cod = Trim(ReplaceStr(Value, "-", ""))
         If Len(Cod) < Len(VFmtCodigoCta(gFmtCodigoCta)) Then   'asumimos que está usando nombre corto
            NombCta = UCase(Trim(Value))
            Cod = ""
         Else
            NombCta = ""
         End If
         
         IdCuenta = GetIdCuenta(NombCta, Cod, DescCta, UltimoNivel)
         
         If IdCuenta = 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
         
         ElseIf UltimoNivel = False Then
            MsgBox1 "No es una cuenta de último nivel.", vbExclamation + vbOKOnly
            Action = vbCancel
         
         Else
            
            Grid.TextMatrix(Row, Col - 1) = IdCuenta
            Value = Format(Cod, gFmtCodigoCta)
            Grid.TextMatrix(Row, Col + 1) = DescCta
            
         End If
         
      Case C_TIPOSOCIO
         Grid.TextMatrix(Row, C_IDTIPOSOCIO) = CbItemData(Grid.CbList(C_TIPOSOCIO))
         
         
   End Select

   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_IDSOCIO, C_UPDATE)
   End If


End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Row > Grid.FixedRows And (Grid.TextMatrix(Row - 1, C_RUT) = "" Or Grid.TextMatrix(Row - 1, C_NOMBRE) = "" Or vFmt(Grid.TextMatrix(Row - 1, C_PJEPART)) = 0 Or Grid.TextMatrix(Row - 1, C_TIPOSOCIO) = "") Then
      MsgBox1 "Debe completar el registro anterior.", vbExclamation
      Exit Sub
   End If
     
   Select Case Col
   
      Case C_RUT
         EdType = FEG_Edit
         Grid.TxBox.MaxLength = 12
   
       Case C_NOMBRE
         EdType = FEG_Edit
         Grid.TxBox.MaxLength = 50
         
       Case C_PJEPART
         EdType = FEG_Edit
         Grid.TxBox.MaxLength = 6
         
       Case C_CANTACCIONES
         EdType = FEG_Edit
         Grid.TxBox.MaxLength = 10
         
       Case C_MONTOSUSCRITO, C_MONTOPAGADO
         EdType = FEG_Edit
         Grid.TxBox.MaxLength = 12
         
      Case C_CODCTAAPORTES, C_CODCTARETIROS
         EdType = FEG_Edit
         
      Case C_TIPOSOCIO
         EdType = FEG_List
         
   End Select
  
   If Row = Grid.rows - 1 Then
      Grid.rows = Grid.rows + 1
   End If
   
End Sub

Private Sub Grid_DblClick()
   Dim FrmPlan As FrmPlanCuentas
   Dim DescCta As String
   Dim CodCta As String
   Dim NombCuenta As String
   Dim Row As Integer, Col As Integer
   Dim IdCuenta As Long
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col = C_CTAAPORTES Or Col = C_CTARETIROS Then

      Set FrmPlan = New FrmPlanCuentas
   
      If FrmPlan.FSelect(IdCuenta, CodCta, DescCta, NombCuenta, True) = vbOK Then
         If DescCta <> "" Then
            
            Grid.TextMatrix(Row, Col - 2) = IdCuenta
            Grid.TextMatrix(Row, Col - 1) = Format(CodCta, gFmtCodigoCta)
            Grid.TextMatrix(Row, Col) = DescCta
               
            Call FGrModRow(Grid, Row, FGR_U, C_IDSOCIO, C_UPDATE)
            
        End If
   
      End If
      Set FrmPlan = Nothing

   End If
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   Select Case Grid.Col
      
      Case C_RUT
         Call KeyRut(KeyAscii)
         
      Case C_NOMBRE
         Call KeyName(KeyAscii)
      
      Case C_PJEPART
         Call KeyDecPos(KeyAscii)
      
      Case C_CANTACCIONES
         Call KeyNumPos(KeyAscii)
      
      Case C_MONTOSUSCRITO, C_MONTOPAGADO
         Call KeyNumPos(KeyAscii)
         
      Case C_CODCTAAPORTES, C_CODCTARETIROS
         Call KeyUpper(KeyAscii)
         
   End Select
         
End Sub

Private Sub Op_TipoContrib_Click(Index As Integer)
   Dim i As Integer

   If lInLoad Then
      Exit Sub
   End If
   
   lTipoContrib = 0
   For i = 1 To MAX_CONTRIB
      If Op_TipoContrib(i) = True Then
         lTipoContrib = i
         Exit For
      End If
   Next i

'   Call EnabFranq
   
   Call SetupTipoContrib
    
End Sub

Private Sub Tx_CodActEcon_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_CodActEcon_LostFocus()
   Call FindCbActEcon
   
End Sub
Private Sub Tx_RUTContador_KeyPress(KeyAscii As Integer)
    Call KeyCID(KeyAscii)
End Sub

Private Sub Tx_RUTContador_LostFocus()
    If Tx_RutContador = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RutContador) Then
      Tx_RutContador.SetFocus
      Exit Sub
      
   End If
      
   MousePointer = vbHourglass
      
   Tx_RutContador = FmtCID(vFmtCID(Tx_RutContador))
   MousePointer = vbDefault
   
End Sub

Private Sub Tx_RUTRep_KeyPress(Index As Integer, KeyAscii As Integer)
    Call KeyCID(KeyAscii)
End Sub

Private Sub Tx_RUTRep_LostFocus(Index As Integer)
    If Tx_RUTRep(Index) = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_RUTRep(Index)) Then
      Tx_RUTRep(Index).SetFocus
      Exit Sub
      
   End If
      
   MousePointer = vbHourglass
      
   Tx_RUTRep(Index) = FmtCID(vFmtCID(Tx_RUTRep(Index)))
   MousePointer = vbDefault
   
End Sub
Private Sub FillCombosFrm()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Codigo As String
   Dim MrkAnt As String
    
   'ACTIVIDAD ECONOMICA
   Set lcbCodActiv = New ClsCombo
   Call lcbCodActiv.SetControl(Cb_ActEcon)
   
   Q1 = "SELECT Descrip, Codigo, Version FROM CodActiv WHERE Version > 2"
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
     ' Cb_ActEcon.AddItem vFld(Rs("Codigo")) & "  " & vFld(Rs("Descrip"), True)
     ' Cb_ActEcon.ItemData(Cb_ActEcon.NewIndex) = Val(vFld(Rs("Codigo")))
      
      '*** PS
'      If vFld(Rs("Version")) = 1 Then
'         MrkAnt = " ! "
'      Else
'         MrkAnt = "   "
'      End If
      MrkAnt = " - "
      lcbCodActiv.AddItem vFld(Rs("Codigo")) & MrkAnt & vFld(Rs("Descrip"), True)
      lcbCodActiv.ItemData(lcbCodActiv.NewIndex) = vFld(Rs("Codigo")) 'Val(vFld(Rs("Codigo")))
      lcbCodActiv.List2(lcbCodActiv.NewIndex) = vFld(Rs("Version"))
     
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
  
   'COMBO REGION
   Call FillRegion(Cb_Region)
   
   Cb_Region.ListIndex = 0
   
   'COMUNA POSTAL, SE MUESTRAN TODAS LAS COMUNAS QU EXISTEN
   Q1 = "SELECT Comuna, id FROM Regiones"
   Q1 = Q1 & " ORDER BY Comuna"
   Cb_ComPostal.AddItem "< Ninguna >"
   Cb_ComPostal.ItemData(Cb_ComPostal.NewIndex) = 0
   Call FillCombo(Cb_ComPostal, DbMain, Q1, -1, True)
   
End Sub
Private Sub FindCbActEcon()
   If Tx_CodActEcon <> "" Then
      'Call SelItem(Cb_ActEcon, Right("00000" & Tx_CodActEcon, 5))
    '  Call SelItem(Cb_ActEcon, Val(Tx_CodActEcon))   'franca
      Call lcbCodActiv.SelItem(Trim(Tx_CodActEcon))
      If lcbCodActiv.ListIndex = -1 Then
         MsgBox1 "Código actividad económica no existe", vbExclamation
         Tx_CodActEcon = ""
         Tx_CodActEcon.SetFocus
         
      End If
   End If
End Sub
Private Sub LoadAll(Optional ByVal TblEmpresa As String = "", Optional ByVal IdEmpresa As Long = 0, Optional ByVal Ano As Long = 0)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Totales(NCOLS) As Double
   Dim Tbl As TableDef
   Dim Mcont As Long
   Dim existe As Boolean
   existe = False
   
   '3189008
   Dim vRutRepLegal1 As String
   Dim vNomRepLegal1 As String
   
   '3189008
   
   
   Tx_Rut = FmtCID(gEmpresa.Rut)
   Tx_NombreCorto = gEmpresa.NombreCorto
      
   Q1 = "SELECT RazonSocial, ApMaterno, Nombre, Calle, Numero,"
   Q1 = Q1 & " EMail, Web, Dpto, Telefonos, Fax, Region, Comuna, Ciudad,"
   Q1 = Q1 & " ActEconom, DomPostal, ComunaPostal, RepConjunta, RutRepLegal1,"
   Q1 = Q1 & " RepLegal1, RutRepLegal2, RepLegal2, CodActEconom, "
   Q1 = Q1 & " Giro, RutContador, Contador, FechaConstitucion, "
   Q1 = Q1 & " FechaInicioAct, "
   Q1 = Q1 & " TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, "
   Q1 = Q1 & " Franq14ter, Franq14quater, FranqRentaAtribuida, FranqSemiIntegrado, ObligaLibComprasVentas, "
   Q1 = Q1 & " FranqSocProfPrimCat, FranqSocProfSegCat, Franq14ASemiIntegrado, "
   Q1 = Q1 & " FranqProPymeGeneral, FranqProPymeTransp, FranqRentasPresuntas, "
   Q1 = Q1 & " FranqRentaEfectiva, FranqOtro, FranqNoSujetoArt14, CodArea, Celular, Villa "
   
   If gDbType = SQL_ACCESS Then
   
     Set Tbl = DbMain.TableDefs(IIf(TblEmpresa = "", "Empresa", TblEmpresa))
    
     For Mcont = 0 To Tbl.Fields.Count - 1
      If Tbl.Fields(Mcont).Name = "CodArea" Then
       existe = True
       Q1 = Q1 & ", CodArea, Celular, Villa "
      End If
     Next Mcont
   
   Else
   
        Q2 = "SELECT COUNT(*) AS TRAE "
        Q2 = Q2 & " FROM INFORMATION_SCHEMA.COLUMNS "
        Q2 = Q2 & " WHERE COLUMN_NAME = 'CodArea' AND TABLE_NAME = '" & IIf(TblEmpresa = "", "Empresa", TblEmpresa) & "' "
        Set Rs = OpenRs(DbMain, Q2)
        
        If Val(Rs("TRAE")) > 0 Then
            existe = True
            Q1 = Q1 & ", CodArea, Celular, Villa "
        End If
        Call CloseRs(Rs)
   
   End If
  
   If TblEmpresa = "" Then
      Q1 = Q1 & " FROM Empresa "
   Else
      Q1 = Q1 & " FROM " & TblEmpresa
   End If
   
   If IdEmpresa = 0 Then
      Q1 = Q1 & " WHERE id = " & gEmpresa.id
   ElseIf IdEmpresa > 0 Then
      Q1 = Q1 & " WHERE id = " & IdEmpresa
   End If
   
   If Ano = 0 Then
      Q1 = Q1 & " AND Ano = " & gEmpresa.Ano
   ElseIf Ano > 0 Then
      Q1 = Q1 & " AND Ano = " & Ano
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      Tx_RazonSocial = vFld(Rs("RazonSocial"), True)
      Tx_Nombre = vFld(Rs("Nombre"), True)
      Tx_ApMaterno = vFld(Rs("ApMaterno"), True)
      Tx_Calle = vFld(Rs("Calle"), True)
      Tx_Numero = vFld(Rs("Numero"))
      Tx_email = vFld(Rs("Email"), True)
      Tx_Web = vFld(Rs("Web"), True)
      Tx_Dpto = vFld(Rs("Dpto"))
      Tx_Telefonos = vFld(Rs("Telefonos"))
      Tx_Fax = vFld(Rs("Fax"))
      Tx_Ciudad = vFld(Rs("Ciudad"))
      Tx_DirPostal = vFld(Rs("DomPostal"))
      Call SetTxDate(Tx_FConstit, vFld(Rs("FechaConstitucion")))
      Call SetTxDate(Tx_FInicioAct, vFld(Rs("FechaInicioAct")))
      Tx_CodActEcon = vFld(Rs("CodActEconom"))
      Ch_RepConjunta = Abs(vFld(Rs("RepConjunta")) <> 0)
      If vFld(Rs("RutRepLegal1")) <> "" Then
         If vFld(Rs("RutRepLegal1")) = "0" Then
            Tx_RUTRep(0) = ""
         Else
            Tx_RUTRep(0) = FmtCID(vFld(Rs("RutRepLegal1")))
            '3189008
            vRutRepLegal1 = FmtCID(vFld(Rs("RutRepLegal1")))
            '3189008
         End If
      End If
      Tx_NombreRep(0) = vFld(Rs("RepLegal1"), True)
      '3189008
      vNomRepLegal1 = vFld(Rs("RepLegal1"), True)
      '3189008
      If vFld(Rs("RutRepLegal2")) <> "" Then
         If vFld(Rs("RutRepLegal2")) = "0" Then
            Tx_RUTRep(1) = ""
         Else
            Tx_RUTRep(1) = FmtCID(vFld(Rs("RutRepLegal2")))
         End If
      End If
      Tx_NombreRep(1) = vFld(Rs("RepLegal2"), True)
      Tx_Contador = vFld(Rs("Contador"), True)
      If vFld(Rs("RutContador")) <> "" Then
         Tx_RutContador = FmtCID(vFld(Rs("RutContador")))
      End If
      Tx_Giro = vFld(Rs("Giro"), True)
      
      Me.Txt_VillaOPob = vFld(Rs("Villa"), True)
      Me.Txt_CodArea = vFld(Rs("CodArea"), True)
      Me.Txt_Celular = vFld(Rs("Celular"), True)
      
      Call SelItem(Cb_Region, vFld(Rs("Region")))
      Call SelItem(Cb_Comuna, vFld(Rs("Comuna")))
      
      'PS Ocupo siempre el CodActEconom, el Otro estaba de antes y ya no es necesario
      If vFld(Rs("CodActEconom"), True) <> "" Then
         Call lcbCodActiv.SelItem(Right("000000" & vFld(Rs("CodActEconom")), 6))
      End If
      
      Call SelItem(Cb_ComPostal, vFld(Rs("ComunaPostal")))
      
      lTipoContrib = vFld(Rs("TipoContrib"))
      
      If lTipoContrib > 0 Then
      
         Op_TipoContrib(lTipoContrib) = True
                  
'         Op_TrBolsaNo = True
'
'         If TipoContrib = CONTRIB_SAABIERTA Then
'            If vFld(Rs("TransaBolsa")) <> 0 Then
'               Op_TrBolsaSi = True
'            End If
'         Else
'            Fr_TrBolsa.Enabled = False
'         End If
'
      Else
         Op_TrBolsaNo = True
         Fr_TrBolsa.Enabled = False

      End If
      
      Ch_Franquicia(FRANQ_14BIS) = Abs(vFld(Rs("Franq14bis")))
      Ch_Franquicia(FRANQ_14QUATER) = Abs(vFld(Rs("Franq14quater")))
      Ch_Franquicia(FRANQ_RENTAATRIB) = Abs(vFld(Rs("FranqRentaAtribuida")))
      Ch_Franquicia(FRANQ_SEMIINTEGRADO) = Abs(vFld(Rs("FranqSemiIntegrado")))
      Ch_Franquicia(FRANQ_14TER) = Abs(vFld(Rs("Franq14ter")))
      
      Ch_ObligaLibComprasVentas = Abs(vFld(Rs("ObligaLibComprasVentas")))
   
      
      Ch_Franquicia(FRANQ_LEY18392) = Abs(vFld(Rs("FranqLey18392")))
      Ch_Franquicia(FRANQ_DL600) = Abs(vFld(Rs("FranqDL600")))
      Ch_Franquicia(FRANQ_DL701) = Abs(vFld(Rs("FranqDL701")))
      Ch_Franquicia(FRANQ_DS341) = Abs(vFld(Rs("FranqDS341")))
      Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) = Abs(vFld(Rs("FranqSocProfPrimCat")))
      Ch_Franquicia(FRANQ_SOCPROFSEGCAT) = Abs(vFld(Rs("FranqSocProfSegCat")))
      
      Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) = Abs(vFld(Rs("Franq14ASemiIntegrado")))
      Ch_Franquicia(FRANQ_PROPYMEGENERAL) = Abs(vFld(Rs("FranqProPymeGeneral")))
      Ch_Franquicia(FRANQ_PROPYMETRANSP) = Abs(vFld(Rs("FranqProPymeTransp")))
      Ch_Franquicia(FRANQ_RENTASPRESUNTAS) = Abs(vFld(Rs("FranqRentasPresuntas")))
      Ch_Franquicia(FRANQ_RENTAEFECTIVA) = Abs(vFld(Rs("FranqRentaEfectiva")))
      Ch_Franquicia(FRANQ_OTRO) = Abs(vFld(Rs("FranqOtro")))
      Ch_Franquicia(FRANQ_NOSUJETOART14) = Abs(vFld(Rs("FranqNoSujetoArt14")))
           
      If existe Then
        Me.Txt_CodArea = vFld(Rs("CodArea"))
        Me.Txt_Celular = vFld(Rs("Celular"))
        Me.Txt_VillaOPob = vFld(Rs("Villa"))
      End If
'       Call EnabFranq
      Call SetupTipoContrib

      
   End If
   Call CloseRs(Rs)
   
   If lGetFromHR Then   'la empresa está en HR y ya tenemos los datos en la estructura gEmpHR (esto sólo debe ser True si es la primera vez que se abre un año para esta empresa y hay datos en HR de esta empresa)
      Call FillFromEmpHR   'sólo llena los campos que están vacíos
   End If
   
   '3189008
   If Len(vRutRepLegal1) > 0 Then
     Tx_RUTRep(0) = vRutRepLegal1
   End If
   
   If Len(vNomRepLegal1) > 0 Then
     Tx_NombreRep(0) = vNomRepLegal1
   End If
   '3189008
   
   
            
   
   
   'ahora los socios
   
   Q1 = "SELECT IdSocio, RUT, Nombre, CantAcciones, PjePart, MontoSuscrito, MontoPagado, IdCuentaAportes, IdCuentaRetiros, IdTipoSocio "
   Q1 = Q1 & " FROM Socios "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Nombre"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDSOCIO) = vFld(Rs("IdSocio"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_PJEPART) = Format(vFld(Rs("PjePart")), DBLFMT2)
      Grid.TextMatrix(i, C_CANTACCIONES) = Format(vFld(Rs("CantAcciones")), NUMFMT)
      Grid.TextMatrix(i, C_MONTOSUSCRITO) = Format(vFld(Rs("MontoSuscrito")), NUMFMT)
      Grid.TextMatrix(i, C_MONTOPAGADO) = Format(vFld(Rs("MontoPagado")), NUMFMT)
      Grid.TextMatrix(i, C_IDCTAAPORTES) = vFld(Rs("IdCuentaAportes"))
      Grid.TextMatrix(i, C_CODCTAAPORTES) = GetCodCuenta(vFld(Rs("IdCuentaAportes")))
      Grid.TextMatrix(i, C_CTAAPORTES) = GetDescCuenta(vFld(Rs("IdCuentaAportes")))
      Grid.TextMatrix(i, C_IDCTARETIROS) = vFld(Rs("IdCuentaRetiros"))
      Grid.TextMatrix(i, C_CODCTARETIROS) = GetCodCuenta(vFld(Rs("IdCuentaRetiros")))
      Grid.TextMatrix(i, C_CTARETIROS) = GetDescCuenta(vFld(Rs("IdCuentaRetiros")))
      Grid.TextMatrix(i, C_IDTIPOSOCIO) = vFld(Rs("IdTipoSocio"))
      Grid.TextMatrix(i, C_TIPOSOCIO) = gTipoSocio(vFld(Rs("IdTipoSocio")))
      
      Totales(C_PJEPART) = Totales(C_PJEPART) + vFld(Rs("PjePart"))
      Totales(C_MONTOSUSCRITO) = Totales(C_MONTOSUSCRITO) + vFld(Rs("MontoSuscrito"))
      Totales(C_MONTOPAGADO) = Totales(C_MONTOPAGADO) + vFld(Rs("MontoPagado"))
      Totales(C_CANTACCIONES) = Totales(C_CANTACCIONES) + vFld(Rs("CantAcciones"))
      
      
      i = i + 1
      Rs.MoveNext
   Loop
   
  
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_RUT
   
   Grid.FlxGrid.Redraw = True
   
   GridTot.TextMatrix(0, C_PJEPART) = Format(Totales(C_PJEPART), DBLFMT2)
   GridTot.TextMatrix(0, C_MONTOSUSCRITO) = Format(Totales(C_MONTOSUSCRITO), NUMFMT)
   GridTot.TextMatrix(0, C_MONTOPAGADO) = Format(Totales(C_MONTOPAGADO), NUMFMT)
   GridTot.TextMatrix(0, C_CANTACCIONES) = Format(Totales(C_CANTACCIONES), NUMFMT)
   
    
   
End Sub
Private Function valida() As Boolean
   Dim i As Integer, r As Integer, Rc As Integer
   
   valida = False
   
   If Trim(Tx_RazonSocial) = "" Then
      MsgBox1 "No se ha ingresado la razón social de la empresa.", vbExclamation
      Exit Function
   End If
   
   If Len(Trim(Tx_RazonSocial)) > 80 Then
      If MsgBox1("Si la Razón Social es muy larga, puede tener problemas al imprimir las hojas foliadas." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
         Exit Function
      End If
   End If
   
   If Trim(Cb_Comuna) <> "" And InStr(Cb_Comuna, "(*)") <> 0 Then
      MsgBox1 "La comuna seleccionada está descontinuada por el SII.", vbExclamation
      Exit Function
   End If
   

   If Trim(Tx_email) <> "" And ValidEmail(Tx_email) = False Then
      MsgBox1 "E-Mail inválido.", vbExclamation
      Exit Function
   End If
      
   For i = 0 To 1
   
      If Trim(Tx_RUTRep(i)) <> "" Then
         If vFmtCID(Tx_RUTRep(i)) > 50000000 Then   'tiene que ser personas naturales
            MsgBox1 "El RUT del representante legal debe corresponder a una persona natural."
            Exit Function
         End If
      End If
   Next i
   
   If Trim(Tx_RUTRep(0)) = "0-0" Then
      Tx_RUTRep(0) = ""
   End If
   If Trim(Tx_RUTRep(1)) = "0-0" Then
      Tx_RUTRep(1) = ""
   End If
   
   If Trim(Tx_RUTRep(0)) <> "" And Trim(Tx_RUTRep(1)) <> "" And vFmtCID(Tx_RUTRep(0)) = vFmtCID(Tx_RUTRep(1)) Then   'tiene que ser personas naturales
      MsgBox1 "Los RUTs de los representantes legales son iguales.", vbExclamation
      Exit Function
   End If
   
   'PS
   If Val(lcbCodActiv.Matrix(2)) = VERSION_1 And gEmpresa.Ano = 2005 Then
      If MsgBox1("¡ ADVERTENCIA ! " & vbNewLine & vbNewLine & "Código de Actividad Económica descontinuado." & vbNewLine & vbNewLine & "¿ Desea continuar ?", vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
         Exit Function
      End If
   End If
   
   If lTipoContrib = 0 Then
      MsgBox1 "Debe seleccionar el Tipo de Contribuyente.", vbExclamation
      Tab1.Tab = TAB_TCONTRIB
      Exit Function
   End If
   
'   If Trim(Val(Me.Txt_CodArea)) < 1 Then
'      MsgBox1 "Favor ingresar el codigo de Area.", vbExclamation
'      Exit Function
'   End If
      
   
   If Ch_Franquicia(FRANQ_RENTAATRIB) <> 0 Or Ch_Franquicia(FRANQ_SEMIINTEGRADO) <> 0 Then
      If Grid.TextMatrix(Grid.FixedRows, C_RUT) = "" Then
         Rc = MsgBox1("ATENCIÓN:" & vbCrLf & vbCrLf & "Recuerde que debe completar el listado de Propietarios y Socios si desea generar el archivo de Retiros y Dividendos" & vbCrLf & vbCrLf & "¿Desea ingresarlos ahora?", vbQuestion + vbYesNoCancel + vbDefaultButton1)
         If Rc = vbCancel Then
            Exit Function
         ElseIf Rc = vbYes Then
            Tab1.Tab = TAB_SOCIOS
            Exit Function
         End If
      End If
   End If
   
   
   'ahora validamos los socios
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_RUT) = "" Then
         Exit For
      End If
   
      If Not ValidCID(Grid.TextMatrix(i, C_RUT)) Then
         MsgBox1 "RUT inválido", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If Trim(Grid.TextMatrix(i, C_NOMBRE)) = "" Then
         MsgBox1 "Falta ingresar el nombre completo", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If (lTipoContrib = CONTRIB_SAABIERTA Or lTipoContrib = CONTRIB_SACERRADA Or lTipoContrib = CONTRIB_SPORACCION) And vFmt(Grid.TextMatrix(i, C_CANTACCIONES)) = 0 Then
         MsgBox1 "Falta ingresar la cantidad de acciones", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If

      
      If vFmt(Grid.TextMatrix(i, C_PJEPART)) = 0 Then
         MsgBox1 "Falta ingresar porcentaje de participación", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If vFmt(Grid.TextMatrix(i, C_MONTOPAGADO)) > vFmt(Grid.TextMatrix(i, C_MONTOSUSCRITO)) Then
         MsgBox1 "El monto pagado reajustado debe ser menor o igual al monto suscrito reajustado", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If vFmt(Grid.TextMatrix(i, C_IDCTAAPORTES)) = 0 Then
         MsgBox1 "Falta ingresar la cuenta contable para los Aportes", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If vFmt(Grid.TextMatrix(i, C_IDCTARETIROS)) = 0 Then
         MsgBox1 "Falta ingresar la cuenta contable para los Retiros", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
      
      If vFmt(Grid.TextMatrix(i, C_IDTIPOSOCIO)) = 0 Then
         MsgBox1 "Falta ingresar el tipo de socio", vbExclamation
         Call FGrSelRow(Grid, i)
         Exit Function
      End If
         
         
      For r = Grid.FixedRows To Grid.rows - 1
     
         If Grid.TextMatrix(r, C_RUT) = "" Then
            Exit For
         End If
         
        If Grid.RowHeight(r) > 0 And (Grid.TextMatrix(r, C_UPDATE) <> "" Or Grid.TextMatrix(r, C_IDSOCIO) <> "") Then
            If r <> i And (StrComp(Grid.TextMatrix(i, C_RUT), Grid.TextMatrix(r, C_RUT), vbTextCompare) = 0 Or StrComp(Grid.TextMatrix(i, C_NOMBRE), Grid.TextMatrix(r, C_NOMBRE), vbTextCompare) = 0) Then
               MsgBox1 "Este Socio o Propietario ya existe.", vbExclamation
               Call FGrSelRow(Grid, i)
               Exit Function
            End If
            
            If r <> i And vFmt(Grid.TextMatrix(i, C_IDCTARETIROS)) = vFmt(Grid.TextMatrix(r, C_IDCTARETIROS)) Then
               MsgBox1 "Esta cuenta de Dividendos/Retiros ya está asignada a otro socio.", vbExclamation
               Call FGrSelRow(Grid, i)
               Exit Function
            End If
            
         End If
      Next r
      
   Next i
   
   If Grid.TextMatrix(Grid.FixedRows, C_RUT) <> "" And vFmt(GridTot.TextMatrix(0, C_PJEPART)) <> 100 Then
      MsgBox1 "La suma de los porcentajes de participación debe ser 100%.", vbExclamation
      Exit Function
   End If
   
   If Ch_Franquicia(FRANQ_14TER) + Ch_Franquicia(FRANQ_SEMIINTEGRADO) + Ch_Franquicia(FRANQ_RENTAATRIB) + Ch_Franquicia(FRANQ_14BIS) + Ch_Franquicia(FRANQ_14QUATER) > 1 Then
      MsgBox1 "Sólo puede elegir un régimen para la empresa.", vbExclamation
      Exit Function
   End If
   
   If Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) + Ch_Franquicia(FRANQ_PROPYMEGENERAL) + Ch_Franquicia(FRANQ_PROPYMETRANSP) + Ch_Franquicia(FRANQ_RENTASPRESUNTAS) + Ch_Franquicia(FRANQ_RENTAEFECTIVA) + Ch_Franquicia(FRANQ_OTRO) + Ch_Franquicia(FRANQ_NOSUJETOART14) > 1 Then
      MsgBox1 "Sólo puede elegir un régimen para la empresa.", vbExclamation
      Exit Function
   End If
   
   valida = True
   
End Function

Private Sub SaveAll()
   Dim Q1 As String
   Dim i As Integer
   Dim CodActEcono As String
   Dim franqSele As Double
  
   ' TEMA 4 2738156
   If CBool(Ch_Franquicia(FRANQ_PROPYMEGENERAL)) Then
   
       If FranqActual <> FRANQ_PROPYMEGENERAL Then
       ' tema 4 2738156
        Q1 = "UPDATE CUENTAS Set CodF22_14Ter = 0 , Atrib8 = 0 WHERE ANO = " & gEmpresa.Ano
   
        Call ExecSQL(DbMain, Q1)
       End If
   ElseIf CBool(Ch_Franquicia(FRANQ_PROPYMETRANSP)) Then
   
       If FranqActual <> FRANQ_PROPYMETRANSP Then
        Q1 = " UPDATE CUENTAS Set CodF22_14Ter = 0 , Atrib8 = 0 WHERE ANO = " & gEmpresa.Ano
   
        Call ExecSQL(DbMain, Q1)
       End If
    End If
   
  
   ' fin TEMA 4 2738156
   
  ' FRANQ_PROPYMEGENERAL
  ' gEmpresa.ProPymeTransp = CInt(Ch_Franquicia(FRANQ_PROPYMETRANSP) <> 0)
   
         
   Q1 = "UPDATE Empresa SET "
   Q1 = Q1 & "RazonSocial='" & ParaSQL(Tx_RazonSocial) & "'"
   Q1 = Q1 & ", Nombre='" & ParaSQL(Tx_Nombre) & "'"
   Q1 = Q1 & ", ApMaterno='" & ParaSQL(Tx_ApMaterno) & "'"
   Q1 = Q1 & ", Calle='" & ParaSQL(Tx_Calle) & "'"
   Q1 = Q1 & ", Numero='" & ParaSQL(Tx_Numero) & "'"
   Q1 = Q1 & ", EMail='" & ParaSQL(Tx_email) & "'"
   Q1 = Q1 & ", Dpto='" & ParaSQL(Tx_Dpto) & "'"
   Q1 = Q1 & ", Telefonos='" & ParaSQL(Tx_Telefonos) & "'"
   Q1 = Q1 & ", Fax='" & ParaSQL(Tx_Fax) & "'"
   Q1 = Q1 & ", Ciudad='" & ParaSQL(Tx_Ciudad) & "'"
   Q1 = Q1 & ", DomPostal='" & ParaSQL(Tx_DirPostal) & "'"
   Q1 = Q1 & ", ComunaPostal=" & ItemData(Cb_ComPostal)
   Q1 = Q1 & ", Web='" & ParaSQL(Tx_Web) & "'"
   'PS, elige una opción en la combo y borra el codigo
   If Trim(Tx_CodActEcon) = "" And lcbCodActiv.ListIndex > 0 Then
      Q1 = Q1 & ", CodActEconom='" & ParaSQL(lcbCodActiv.ItemData) & "'"
      CodActEcono = lcbCodActiv.ItemData
   Else
      Q1 = Q1 & ", CodActEconom='" & ParaSQL(Tx_CodActEcon) & "'"
      CodActEcono = Trim(Tx_CodActEcon)
   End If
'   Q1 = Q1 & ", ActEconom =" & lcbCodActiv.ItemData  'ItemData(Cb_ActEcon) 'Este ya no se usa, porque el Codigo es String
   Q1 = Q1 & ", ActEconom = 0"
   Q1 = Q1 & ", FechaConstitucion=" & GetTxDate(Tx_FConstit)
   Q1 = Q1 & ", FechaInicioAct=" & GetTxDate(Tx_FInicioAct)
   Q1 = Q1 & ", Region=" & ItemData(Cb_Region)
   Q1 = Q1 & ", Comuna=" & ItemData(Cb_Comuna)
   Q1 = Q1 & ", RepConjunta=" & IIf(Ch_RepConjunta <> 0, 1, 0)
   Q1 = Q1 & ", RutRepLegal1='" & IIf(Trim(Tx_RUTRep(0)) <> "", vFmtCID(Tx_RUTRep(0)), " ") & "'"
   Q1 = Q1 & ", RutRepLegal2='" & IIf(Trim(Tx_RUTRep(1)) <> "", vFmtCID(Tx_RUTRep(1)), " ") & "'"
   Q1 = Q1 & ", RepLegal1='" & ParaSQL(Tx_NombreRep(0)) & "'"
   Q1 = Q1 & ", RepLegal2='" & ParaSQL(Tx_NombreRep(1)) & "'"
   Q1 = Q1 & ", Giro='" & ParaSQL(Tx_Giro) & "'"
   Q1 = Q1 & ", Contador='" & ParaSQL(Tx_Contador) & "'"
   Q1 = Q1 & ", RutContador='" & IIf(Trim(Tx_RutContador) <> "", vFmtCID(Tx_RutContador), " ") & "'"
   
        
   Q1 = Q1 & ", TipoContrib=" & lTipoContrib
'   Q1 = Q1 & ", TContribFUT=" & TipoContrib
   Q1 = Q1 & ", TransaBolsa=" & CInt(Op_TrBolsaSi = True)
   Q1 = Q1 & ", Franq14bis=" & CInt(Ch_Franquicia(FRANQ_14BIS) <> 0)
   Q1 = Q1 & ", FranqLey18392=" & CInt(Ch_Franquicia(FRANQ_LEY18392) <> 0)
   Q1 = Q1 & ", FranqDL600=" & CInt(Ch_Franquicia(FRANQ_DL600) <> 0)
   Q1 = Q1 & ", FranqDL701=" & CInt(Ch_Franquicia(FRANQ_DL701) <> 0)
   Q1 = Q1 & ", FranqDS341=" & CInt(Ch_Franquicia(FRANQ_DS341) <> 0)
   Q1 = Q1 & ", Franq14ter=" & CInt(Ch_Franquicia(FRANQ_14TER) <> 0)
   Q1 = Q1 & ", Franq14quater=" & CInt(Ch_Franquicia(FRANQ_14QUATER) <> 0)
   Q1 = Q1 & ", FranqRentaAtribuida=" & CInt(Ch_Franquicia(FRANQ_RENTAATRIB) <> 0)
   Q1 = Q1 & ", FranqSemiIntegrado=" & CInt(Ch_Franquicia(FRANQ_SEMIINTEGRADO) <> 0)
   Q1 = Q1 & ", FranqSocProfPrimCat=" & CInt(Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) <> 0)
   Q1 = Q1 & ", FranqSocProfSegCat=" & CInt(Ch_Franquicia(FRANQ_SOCPROFSEGCAT) <> 0)
   
   Q1 = Q1 & ", Franq14ASemiIntegrado=" & CInt(Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) <> 0)
   Q1 = Q1 & ", FranqProPymeGeneral=" & CInt(Ch_Franquicia(FRANQ_PROPYMEGENERAL) <> 0)
   Q1 = Q1 & ", FranqProPymeTransp=" & CInt(Ch_Franquicia(FRANQ_PROPYMETRANSP) <> 0)
   Q1 = Q1 & ", FranqRentasPresuntas=" & CInt(Ch_Franquicia(FRANQ_RENTASPRESUNTAS) <> 0)
   Q1 = Q1 & ", FranqRentaEfectiva=" & CInt(Ch_Franquicia(FRANQ_RENTAEFECTIVA) <> 0)
   Q1 = Q1 & ", FranqOtro=" & CInt(Ch_Franquicia(FRANQ_OTRO) <> 0)
   Q1 = Q1 & ", FranqNoSujetoArt14=" & CInt(Ch_Franquicia(FRANQ_NOSUJETOART14) <> 0)
   
   Q1 = Q1 & ", ObligaLibComprasVentas=" & CInt(Ch_ObligaLibComprasVentas <> 0)
   Q1 = Q1 & ", CodArea=" & IIf(Me.Txt_CodArea = "", 0, ParaSQL(Me.Txt_CodArea))
   Q1 = Q1 & ", Celular=" & IIf(Me.Txt_Celular = "", 0, ParaSQL(Me.Txt_Celular))
   Q1 = Q1 & ", Villa='" & ParaSQL(Me.Txt_VillaOPob) & "'"
   Q1 = Q1 & " WHERE Id =" & gEmpresa.id & " AND Ano =" & gEmpresa.Ano
            
   Call ExecSQL(DbMain, Q1)
   
   lRc = vbOK
      
   Q1 = "UPDATE ControlEmpresa SET"
   Q1 = Q1 & " RazonSocial= '" & ParaSQL(Tx_RazonSocial) & "'"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)

   gEmpresa.RazonSocial = Trim(Tx_RazonSocial) & " " & Trim(Tx_ApMaterno) & " " & Trim(Tx_Nombre)
   gEmpresa.Direccion = Trim(Tx_Calle) & " " & Trim(Tx_Numero) & " " & Trim(Tx_Dpto)
   gEmpresa.Telefono = Trim(Tx_Telefonos)
   gEmpresa.Region = ItemData(Cb_Region)
   gEmpresa.Comuna = FCase(Cb_Comuna)
   gEmpresa.Ciudad = Trim(Tx_Ciudad)
   gEmpresa.Fax = Trim(Tx_Fax)
   gEmpresa.Giro = Trim(Tx_Giro)
   gEmpresa.CodActEcono = CodActEcono
   gEmpresa.RepConjunta = (Ch_RepConjunta <> 0)
   gEmpresa.RutRepLegal1 = vFmtCID(Tx_RUTRep(0))
   gEmpresa.RutRepLegal2 = vFmtCID(Tx_RUTRep(1))
   gEmpresa.RepLegal1 = ParaSQL(Tx_NombreRep(0))
   gEmpresa.RepLegal2 = ParaSQL(Tx_NombreRep(1))
   gEmpresa.Franq14Ter = CInt(Ch_Franquicia(FRANQ_14TER) <> 0)
   gEmpresa.RentaAtribuida = CInt(Ch_Franquicia(FRANQ_RENTAATRIB) <> 0)
   gEmpresa.SocProfSegCat = CInt(Ch_Franquicia(FRANQ_SOCPROFSEGCAT) <> 0)
   gEmpresa.SemiIntegrado = CInt(Ch_Franquicia(FRANQ_SEMIINTEGRADO) <> 0)
   gEmpresa.ObligaLibComprasVentas = CInt(Ch_ObligaLibComprasVentas <> 0)
   gEmpresa.TipoContrib = lTipoContrib
   
   gEmpresa.R14ASemiIntegrado = CInt(Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) <> 0)
   gEmpresa.ProPymeGeneral = CInt(Ch_Franquicia(FRANQ_PROPYMEGENERAL) <> 0)
   gEmpresa.ProPymeTransp = CInt(Ch_Franquicia(FRANQ_PROPYMETRANSP) <> 0)
   
   If gEmpresa.Ano >= 2020 Then
      'si cambia desde o hacia regimen 14D de algún tipo o 14A, hacemos homologación de campo CodF22_14Ter en tabla Cuentas
      
      Call Homologa_CodF22_14Ter_14D

   End If


   Call SetPrtData
   
   'y ahora los socios
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_RUT) = "" Then
         Exit For
      End If
         
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
         Q1 = "INSERT INTO Socios (RUT, Nombre, CantAcciones, PjePart, MontoSuscrito, MontoPagado, IdCuentaAportes, IdCuentaRetiros, IdTipoSocio, IdEmpresa, Ano, Vigente ) "
         Q1 = Q1 & " VALUES ("
         Q1 = Q1 & " '" & vFmtCID(Grid.TextMatrix(i, C_RUT)) & "'"
         Q1 = Q1 & ",'" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_CANTACCIONES))
         Q1 = Q1 & "," & str(vFmt(Grid.TextMatrix(i, C_PJEPART)))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_MONTOSUSCRITO))
         Q1 = Q1 & "," & vFmt(Grid.TextMatrix(i, C_MONTOPAGADO))
         Q1 = Q1 & "," & Grid.TextMatrix(i, C_IDCTAAPORTES)
         Q1 = Q1 & "," & Grid.TextMatrix(i, C_IDCTARETIROS)
         Q1 = Q1 & "," & Grid.TextMatrix(i, C_IDTIPOSOCIO)
         Q1 = Q1 & "," & gEmpresa.id
         Q1 = Q1 & "," & gEmpresa.Ano
         Q1 = Q1 & ", -1 )"
         
         Call ExecSQL(DbMain, Q1)
      
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
         Q1 = "UPDATE Socios SET "
         Q1 = Q1 & "  RUT='" & vFmtCID(Grid.TextMatrix(i, C_RUT)) & "'"
         Q1 = Q1 & ", Nombre='" & ParaSQL(Grid.TextMatrix(i, C_NOMBRE)) & "'"
         Q1 = Q1 & ", CantAcciones= " & vFmt(Grid.TextMatrix(i, C_CANTACCIONES))
         Q1 = Q1 & ", PjePart= " & str(vFmt(Grid.TextMatrix(i, C_PJEPART)))
         Q1 = Q1 & ", MontoSuscrito= " & vFmt(Grid.TextMatrix(i, C_MONTOSUSCRITO))
         Q1 = Q1 & ", MontoPagado= " & vFmt(Grid.TextMatrix(i, C_MONTOPAGADO))
         Q1 = Q1 & ", IdCuentaAportes= " & Grid.TextMatrix(i, C_IDCTAAPORTES)
         Q1 = Q1 & ", IdCuentaRetiros= " & Grid.TextMatrix(i, C_IDCTARETIROS)
         Q1 = Q1 & ", IdTipoSocio= " & Grid.TextMatrix(i, C_IDTIPOSOCIO)
         
         Q1 = Q1 & " WHERE IdSocio=" & Grid.TextMatrix(i, C_IDSOCIO)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
      
'         Q1 = "DELETE * FROM Socios "
         Q1 = " WHERE IdSocio = " & Grid.TextMatrix(i, C_IDSOCIO)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'         Call ExecSQL(DbMain, Q1)
      
         Call DeleteSQL(DbMain, "Socios", Q1)
      End If
      
   Next i

   
End Sub
Private Sub Tx_FConstit_GotFocus()
   Call DtGotFocus(Tx_FConstit)
End Sub

Private Sub Tx_FConstit_LostFocus()
   
   If Trim$(Tx_FConstit) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FConstit)
   
End Sub

Private Sub Tx_FConstit_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_FInicioAct_GotFocus()
   Call DtGotFocus(Tx_FInicioAct)
End Sub

Private Sub Tx_FInicioAct_LostFocus()
   
   If Trim$(Tx_FInicioAct) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FInicioAct)
   
End Sub

Private Sub Tx_FInicioAct_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub
'PS
Private Sub La_Url_Click()
   Dim Rc As Long
   Dim Url As String
   
   Url = "http://www.sii.cl/catastro/codigos_economica.htm"
   Rc = Shell(gHtmExt.OpenCmd & " " & Url, vbNormalFocus)

End Sub

Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
   
   WCodCuenta = Me.TextWidth(gFmtCodigoCta) + 300
   WCuenta = 1450
   
   Call FGrSetup(Grid)

   Grid.ColWidth(C_IDSOCIO) = 0
   Grid.ColWidth(C_RUT) = 1150
   Grid.ColWidth(C_NOMBRE) = 2000
   Grid.ColWidth(C_PJEPART) = 600
   Grid.ColWidth(C_MONTOSUSCRITO) = 1200
   Grid.ColWidth(C_MONTOPAGADO) = 1200
   Grid.ColWidth(C_IDCTAAPORTES) = 0
   Grid.ColWidth(C_CODCTAAPORTES) = 0     'WCodCuenta
   Grid.ColWidth(C_CANTACCIONES) = 0
   Grid.ColWidth(C_CTAAPORTES) = WCuenta
   Grid.ColWidth(C_IDCTARETIROS) = 0
   Grid.ColWidth(C_CODCTARETIROS) = 0     'WCodCuenta
   Grid.ColWidth(C_CTARETIROS) = WCuenta
   Grid.ColWidth(C_IDTIPOSOCIO) = 0
   Grid.ColWidth(C_TIPOSOCIO) = 2000
   Grid.ColWidth(C_UPDATE) = 0
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_PJEPART) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOSUSCRITO) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOPAGADO) = flexAlignRightCenter
   Grid.ColAlignment(C_CANTACCIONES) = flexAlignRightCenter
   
   Grid.TextMatrix(1, C_RUT) = "RUT"
   Grid.TextMatrix(1, C_NOMBRE) = "Nombre Completo"
   Grid.TextMatrix(1, C_PJEPART) = "% Part."
   Grid.TextMatrix(0, C_MONTOSUSCRITO) = "Monto Suscrito"
   Grid.TextMatrix(1, C_MONTOSUSCRITO) = "Reajustado $"
   Grid.TextMatrix(0, C_MONTOPAGADO) = "Monto Pagado"
   Grid.TextMatrix(1, C_MONTOPAGADO) = "Reajustado $"
'   Grid.TextMatrix(0, C_CODCTAAPORTES) = "Cód. Cuenta"
'   Grid.TextMatrix(1, C_CODCTAAPORTES) = "Aportes"
   Grid.TextMatrix(0, C_CTAAPORTES) = "Cuenta"
   Grid.TextMatrix(1, C_CTAAPORTES) = "Aportes"
'   Grid.TextMatrix(0, C_CODCTARETIROS) = "Cód. Cuenta"
'   Grid.TextMatrix(1, C_CODCTARETIROS) = "Retiros"
   Grid.TextMatrix(0, C_CTARETIROS) = "Cuenta"
   Grid.TextMatrix(1, C_CTARETIROS) = "Retiros"
   Grid.TextMatrix(1, C_TIPOSOCIO) = "Tipo de Socio"
   
   Call FGrVRows(Grid)
   Call FGrTotales(Grid, GridTot)
   
   For i = 0 To MAX_TIPOSOCIO
      Call CbAddItem(Grid.CbList(C_TIPOSOCIO), gTipoSocio(i), i)
   Next i
   
   
End Sub

Private Sub Grid_SelChange()
   Dim EdType As FlexEdGrid2.FEG2_EdType
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Col = C_CTAAPORTES Then
      Tx_CurrCell = Grid.TextMatrix(Grid.Row, C_CODCTAAPORTES) & " - " & Grid.TextMatrix(Grid.Row, C_CTAAPORTES)
   ElseIf Grid.Col = C_CTARETIROS Then
      Tx_CurrCell = Grid.TextMatrix(Grid.Row, C_CODCTARETIROS) & " - " & Grid.TextMatrix(Grid.Row, C_CTARETIROS)
   Else
      Tx_CurrCell = Grid.TextMatrix(Grid.Row, Grid.Col)
   End If

End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   If MsgBox1("¿Está seguro que desea elimnar este socio o propietario?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Call FGrModRow(Grid, Row, FGR_D, C_IDSOCIO, C_UPDATE)
   Grid.rows = Grid.rows + 1
   Call CalcTotSocios


End Sub
Private Sub CalcTotSocios()
   Dim i As Integer
   Dim Totales(NCOLS) As Double
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_RUT) = "" Then
         Exit For
      End If
      
      If Grid.RowHeight(i) > 0 Then
         Totales(C_PJEPART) = Totales(C_PJEPART) + vFmt(Grid.TextMatrix(i, C_PJEPART))
         Totales(C_MONTOSUSCRITO) = Totales(C_MONTOSUSCRITO) + vFmt(Grid.TextMatrix(i, C_MONTOSUSCRITO))
         Totales(C_MONTOPAGADO) = Totales(C_MONTOPAGADO) + vFmt(Grid.TextMatrix(i, C_MONTOPAGADO))
         Totales(C_CANTACCIONES) = Totales(C_CANTACCIONES) + vFmt(Grid.TextMatrix(i, C_CANTACCIONES))
      End If
      
   Next i
   
   GridTot.TextMatrix(0, C_PJEPART) = Format(Totales(C_PJEPART), DBLFMT2)
   GridTot.TextMatrix(0, C_MONTOSUSCRITO) = Format(Totales(C_MONTOSUSCRITO), NUMFMT)
   GridTot.TextMatrix(0, C_MONTOPAGADO) = Format(Totales(C_MONTOPAGADO), NUMFMT)
   GridTot.TextMatrix(0, C_CANTACCIONES) = Format(Totales(C_CANTACCIONES), NUMFMT)
End Sub
Private Sub SetupPriv()
   Dim i As Integer
   Dim bool As Boolean

   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   

End Sub

Private Sub FillFromEmpHR()

   If gEmprHR.EmpConta.RazonSocial = "" Then
      Exit Sub
   End If

   Tx_RazonSocial = IIf(Tx_RazonSocial = "", gEmprHR.EmpConta.RazonSocial, Tx_RazonSocial)
   Tx_Nombre = IIf(Tx_Nombre = "", gEmprHR.EmpConta.Nombre, Tx_Nombre)
   Tx_ApMaterno = IIf(Tx_ApMaterno = "", gEmprHR.ApMaterno, Tx_ApMaterno)
   Tx_Calle = IIf(Tx_Calle = "", gEmprHR.EmpConta.Direccion, Tx_Calle)
   Tx_Numero = IIf(Tx_Numero = "", gEmprHR.NroCalle, Tx_Numero)
   Tx_Dpto = IIf(Tx_Dpto = "", gEmprHR.NroDepto, Tx_Dpto)
   Tx_email = IIf(Tx_email = "", gEmprHR.EmpConta.email, Tx_email)
   Tx_Telefonos = IIf(Tx_Telefonos = "", gEmprHR.EmpConta.Telefono, Tx_Telefonos)
   Tx_Fax = IIf(Tx_Fax = "", gEmprHR.EmpConta.Fax, Tx_Fax)
   Tx_Ciudad = IIf(Tx_Ciudad = "", gEmprHR.EmpConta.Ciudad, Tx_Ciudad)
   Cb_Region.ListIndex = CbSelItem(Cb_Region, Val(gEmprHR.Region))
   Cb_Comuna.ListIndex = CbFindText(Cb_Comuna, gEmprHR.EmpConta.Comuna)
   Cb_ComPostal.ListIndex = CbFindText(Cb_ComPostal, gEmprHR.ComunaPostal)
   Tx_Giro = IIf(Tx_Giro = "", gEmprHR.EmpConta.Giro, Tx_Giro)
   Me.Txt_Celular = IIf(Txt_Celular = 0 Or Txt_Celular = "", gEmprHR.EmpConta.Celular, Txt_Celular)
   Me.Txt_CodArea = IIf(Txt_CodArea = 0 Or Txt_CodArea = "", gEmprHR.EmpConta.CodArea, Txt_CodArea)
   Me.Txt_VillaOPob = IIf(Txt_VillaOPob = "", gEmprHR.EmpConta.Villa, Txt_VillaOPob)
   lcbCodActiv.ListIndex = lcbCodActiv.FindItem(gEmprHR.EmpConta.CodActEcono)

   Tx_RUTRep(0) = IIf(Tx_RUTRep(0) = "", IIf(gEmprHR.EmpConta.RutRepLegal1 = "", "", FmtCID(gEmprHR.EmpConta.RutRepLegal1)), Tx_RUTRep(0))
   Tx_NombreRep(0) = IIf(Tx_NombreRep(0) = "", gEmprHR.EmpConta.RepLegal1, Tx_NombreRep(0))
   
   Tx_DirPostal = gEmprHR.DirPostal
   Tx_Contador = gEmprHR.NombContador
   Tx_RutContador = gEmprHR.RutContador
   
   If gEmprHR.TipoContrib > 0 Then
      Op_TipoContrib(gEmprHR.TipoContrib) = True
   End If
   
   lTipoContrib = gEmprHR.TipoContrib
   
   If gEmprHR.TransaBolsa Then
      Op_TrBolsaSi = 1
   Else
      Op_TrBolsaNo = 1
   End If
   
   
   If gEmpresa.Ano < 2017 Then
      Ch_Franquicia(FRANQ_14BIS) = gEmprHR.Franquicias(FRANQ_14BIS)
'      Ch_Franquicia(FRANQ_14QUATER) = Abs(vFld(Rs("Franq14quater")))
   Else
      Ch_Franquicia(FRANQ_14BIS).Enabled = 0
      Ch_Franquicia(FRANQ_14QUATER).Enabled = 0
      Ch_Franquicia(FRANQ_RENTAATRIB) = IIf(gEmprHR.EmpConta.RentaAtribuida, 1, 0)
      Ch_Franquicia(FRANQ_SEMIINTEGRADO) = IIf(gEmprHR.EmpConta.SemiIntegrado, 1, 0)
   End If
   
   '14 TER
   If (lTipoContrib = CONTRIB_SPORACCION Or lTipoContrib = CONTRIB_PRIMCAT Or lTipoContrib = CONTRIB_EMPINDIVIDUALEIRL Or lTipoContrib = CONTRIB_EMPINDIVIDUAL) And gEmpresa.Ano > 2014 Then
      Ch_Franquicia(FRANQ_14TER).Enabled = True
      Ch_Franquicia(FRANQ_14TER) = IIf(gEmprHR.Franquicias(FRANQ_14TER), 1, 0)
   Else
      Ch_Franquicia(FRANQ_14TER).Enabled = False
      Ch_Franquicia(FRANQ_14TER) = 0
   End If
   
   Ch_Franquicia(FRANQ_LEY18392) = IIf(gEmprHR.Franquicias(FRANQ_LEY18392), 1, 0)
   Ch_Franquicia(FRANQ_DL600) = IIf(gEmprHR.Franquicias(FRANQ_DL600), 1, 0)
   Ch_Franquicia(FRANQ_DL701) = IIf(gEmprHR.Franquicias(FRANQ_DL701), 1, 0)
   Ch_Franquicia(FRANQ_DS341) = IIf(gEmprHR.Franquicias(FRANQ_DS341), 1, 0)
   
   If Ch_Franquicia(FRANQ_14TER) = 0 Then
      Ch_ObligaLibComprasVentas.Enabled = False
      Ch_ObligaLibComprasVentas = 0
'   Else
'      Ch_ObligaLibComprasVentas = Abs(vFld(Rs("ObligaLibComprasVentas")))
   End If

   If gEmpresa.Ano < 2020 Then
      Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) = 0
      Ch_Franquicia(FRANQ_PROPYMEGENERAL) = 0
      Ch_Franquicia(FRANQ_PROPYMETRANSP) = 0
      Ch_Franquicia(FRANQ_RENTASPRESUNTAS) = 0
      Ch_Franquicia(FRANQ_RENTAEFECTIVA) = 0
      Ch_Franquicia(FRANQ_OTRO) = 0
      Ch_Franquicia(FRANQ_NOSUJETOART14) = 0
   Else
      Ch_Franquicia(FRANQ_14ASEMIINTEGRADO) = IIf(gEmprHR.EmpConta.R14ASemiIntegrado, 1, 0)
      Ch_Franquicia(FRANQ_PROPYMEGENERAL) = IIf(gEmprHR.EmpConta.ProPymeGeneral, 1, 0)
      Ch_Franquicia(FRANQ_PROPYMETRANSP) = IIf(gEmprHR.EmpConta.ProPymeTransp, 1, 0)
      Ch_Franquicia(FRANQ_RENTASPRESUNTAS) = IIf(gEmprHR.EmpConta.RentasPresuntas, 1, 0)
      Ch_Franquicia(FRANQ_RENTAEFECTIVA) = IIf(gEmprHR.EmpConta.RentaEfectiva, 1, 0)
      Ch_Franquicia(FRANQ_OTRO) = IIf(gEmprHR.EmpConta.RegimenOtro, 1, 0)
      Ch_Franquicia(FRANQ_NOSUJETOART14) = IIf(gEmprHR.EmpConta.NoSujetoArt14, 1, 0)
   End If

End Sub

Private Sub EnabFranq()
   Dim i As Integer

   For i = 1 To MAX_FRANQ
      Ch_Franquicia(i).Enabled = True
   Next i
   
   '14 TER
   If Ch_Franquicia(FRANQ_14TER) <> 0 Then
      Ch_ObligaLibComprasVentas.Enabled = True
   Else
      Ch_ObligaLibComprasVentas = 0
      Ch_ObligaLibComprasVentas.Enabled = False
   End If
   

   If lTipoContrib = CONTRIB_SAABIERTA Then
      Fr_TrBolsa.Enabled = True
   Else
      Fr_TrBolsa.Enabled = False
      Op_TrBolsaNo = True
   End If


   If lTipoContrib = CONTRIB_SAABIERTA Or lTipoContrib = CONTRIB_SACERRADA Then
      For i = 1 To MAX_FRANQ
         If i <> FRANQ_SEMIINTEGRADO Then
            Ch_Franquicia(i).Enabled = False
         End If
      Next i
   End If

   If (lTipoContrib = CONTRIB_SPORACCION Or lTipoContrib = CONTRIB_PRIMCAT Or lTipoContrib = CONTRIB_EMPINDIVIDUALEIRL Or lTipoContrib = CONTRIB_EMPINDIVIDUAL) And gEmpresa.Ano > 2014 Then
      Ch_Franquicia(FRANQ_14TER).Enabled = True
   Else
      Ch_Franquicia(FRANQ_14TER).Enabled = False
      Ch_Franquicia(FRANQ_14TER) = 0
   End If
   
   If gEmpresa.Ano < 2017 Then
      Ch_Franquicia(FRANQ_RENTAATRIB) = 0
      Ch_Franquicia(FRANQ_RENTAATRIB).Enabled = False
      Ch_Franquicia(FRANQ_SEMIINTEGRADO) = 0
      Ch_Franquicia(FRANQ_SEMIINTEGRADO).Enabled = False
   Else
      Ch_Franquicia(FRANQ_14QUATER) = 0
      Ch_Franquicia(FRANQ_14QUATER).Enabled = False
      Ch_Franquicia(FRANQ_14BIS) = 0
      Ch_Franquicia(FRANQ_14BIS).Enabled = False
   
   End If

   If lTipoContrib = CONTRIB_SOCPROFESIONAL Then
   
      For i = 1 To MAX_FRANQ
         If Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) <> 0 Then
            If i <> FRANQ_SOCPROFPRIMCAT And i <> FRANQ_SOCPROFSEGCAT And i <> FRANQ_RENTAATRIB And i <> FRANQ_SEMIINTEGRADO And i <> FRANQ_14TER And i <> FRANQ_SOCPROFPRIMCAT Then
               Ch_Franquicia(i) = 0
               Ch_Franquicia(i).Enabled = False
            Else
               Ch_Franquicia(i).Enabled = True
            End If
            
         ElseIf Ch_Franquicia(FRANQ_SOCPROFSEGCAT) <> 0 Then
            If i <> FRANQ_SOCPROFPRIMCAT And i <> FRANQ_SOCPROFSEGCAT Then
               Ch_Franquicia(i) = 0
               Ch_Franquicia(i).Enabled = False
            Else
               Ch_Franquicia(i).Enabled = True
            End If
         End If
      Next i
      
   Else
      Ch_Franquicia(FRANQ_SOCPROFPRIMCAT).Enabled = False
      Ch_Franquicia(FRANQ_SOCPROFPRIMCAT) = 0
      Ch_Franquicia(FRANQ_SOCPROFSEGCAT).Enabled = False
      Ch_Franquicia(FRANQ_SOCPROFSEGCAT) = 0
   End If

End Sub

Private Sub SetupTipoContrib()

   If lTipoContrib = CONTRIB_SAABIERTA Or lTipoContrib = CONTRIB_SACERRADA Or lTipoContrib = CONTRIB_SPORACCION Then
   
      Tab1.TabCaption(TAB_SOCIOS) = "Propietarios y Accionistas"
   
      Grid.TextMatrix(0, C_CTARETIROS) = "Cuenta"
      Grid.TextMatrix(1, C_CTARETIROS) = "Dividendos"
      Grid.TextMatrix(1, C_TIPOSOCIO) = "Tipo de Accionista"
      Grid.TextMatrix(0, C_CANTACCIONES) = "Cant."
      Grid.TextMatrix(1, C_CANTACCIONES) = "Acc."
      Grid.ColWidth(C_CANTACCIONES) = 800
      GridTot.ColWidth(C_CANTACCIONES) = 600
   
   Else
   
      Tab1.TabCaption(TAB_SOCIOS) = "Propietarios y Socios"
   
      Grid.TextMatrix(0, C_CTARETIROS) = "Cuenta"
      Grid.TextMatrix(1, C_CTARETIROS) = "Retiros"
      Grid.TextMatrix(1, C_TIPOSOCIO) = "Tipo de Socio"
      Grid.TextMatrix(1, C_CANTACCIONES) = ""
      Grid.ColWidth(C_CANTACCIONES) = 0
      GridTot.ColWidth(C_CANTACCIONES) = 600
      
   End If
End Sub
Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub



Private Sub Tx_Telefonos_KeyPress(KeyAscii As Integer)
Call KeyNumPos(KeyAscii)
End Sub

Private Sub Txt_Celular_KeyPress(KeyAscii As Integer)
Call KeyNumPos(KeyAscii)
End Sub

Private Sub Txt_CodArea_KeyPress(KeyAscii As Integer)
Call KeyNumPos(KeyAscii)
End Sub
