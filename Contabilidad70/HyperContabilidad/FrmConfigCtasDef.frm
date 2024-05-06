VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConfigCtasDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Cuentas Básicas"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "FrmConfigCtasDef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   360
      Picture         =   "FrmConfigCtasDef.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   600
      TabIndex        =   65
      Top             =   300
      Width           =   600
   End
   Begin TabDlg.SSTab St_tab 
      Height          =   9255
      Left            =   1200
      TabIndex        =   22
      Top             =   300
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   16325
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Impuestos"
      TabPicture(0)   =   "FrmConfigCtasDef.frx":05AF
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame5(1)"
      Tab(0).Control(3)=   "Frame5(0)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Libros"
      TabPicture(1)   =   "FrmConfigCtasDef.frx":05CB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame5(2)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Otros"
      TabPicture(2)   =   "FrmConfigCtasDef.frx":05E7
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Fr_IVAIrrec"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Fr_Ret3Porc"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Fr_ppm"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame10"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame10 
         Caption         =   "Cuentas ODF"
         Height          =   1215
         Left            =   360
         TabIndex        =   89
         Top             =   6360
         Width           =   6855
         Begin VB.TextBox Tx_OdfActivo 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   93
            ToolTipText     =   "Indicar medio de pago de egresos (Caja, Banco, Tarjetas, ect.)"
            Top             =   240
            Width           =   3855
         End
         Begin VB.CommandButton Bt_OdfActivo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":0603
            Style           =   1  'Graphical
            TabIndex        =   92
            ToolTipText     =   "Plan de Cuentas"
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton Bt_OdfPasivo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":09C4
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Plan de Cuentas"
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox Tx_OdfPasivo 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   90
            ToolTipText     =   "Indicar medio de recaudación de dineros (Caja, Banco, Tarjetas, ect.)"
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Activo:"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   95
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo:"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   94
            Top             =   660
            Width           =   525
         End
      End
      Begin VB.Frame Fr_ppm 
         Caption         =   "Cuentas PPM"
         Height          =   1215
         Left            =   360
         TabIndex        =   82
         Top             =   7680
         Width           =   6855
         Begin VB.TextBox Tx_PpmVoluntario 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   86
            ToolTipText     =   "Indicar medio de recaudación de dineros (Caja, Banco, Tarjetas, ect.)"
            Top             =   600
            Width           =   3855
         End
         Begin VB.CommandButton Bt_PpmVolunt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":0D85
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Plan de Cuentas"
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton Bt_PpmOblig 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":1146
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Plan de Cuentas"
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox Tx_PpmObligatorio 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   83
            ToolTipText     =   "Indicar medio de pago de egresos (Caja, Banco, Tarjetas, ect.)"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PPM Voluntario:"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   88
            Top             =   660
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PPM Obligatorio:"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   87
            Top             =   300
            Width           =   1185
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "3% Ret. Centralización Remuneraciones"
         Height          =   855
         Left            =   -74640
         TabIndex        =   79
         Top             =   5880
         Width           =   6915
         Begin VB.CommandButton Bt_Imp3CentraRem 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":1507
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Tx_Imp3CentraRem 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   300
            Width           =   3795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 3% Ret. Remu."
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.Frame Fr_Ret3Porc 
         Caption         =   "Retención  3% Préstamo Tasa Cero Libro de Retenciones"
         Height          =   855
         Left            =   360
         TabIndex        =   19
         Top             =   5340
         Width           =   6915
         Begin VB.CommandButton Bt_Ret3Porc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":18C8
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Tx_Ret3Porc 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   300
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta de Retención 3%"
            Height          =   255
            Index           =   19
            Left            =   180
            TabIndex        =   78
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Fr_IVAIrrec 
         Caption         =   "IVA Irrecuperable (para Cálculo de Proporcionalidad)"
         Height          =   855
         Left            =   360
         TabIndex        =   62
         Top             =   4380
         Width           =   6915
         Begin VB.TextBox Tx_IVAIrrec 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   300
            Width           =   3855
         End
         Begin VB.CommandButton Bt_IVAIrrec 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":1C89
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta IVA Irrecuperable"
            Height          =   255
            Index           =   17
            Left            =   180
            TabIndex        =   64
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Remanente IVA Año Antrerior"
         Height          =   855
         Left            =   360
         TabIndex        =   59
         Top             =   3420
         Width           =   6915
         Begin VB.CommandButton Bt_CredIVA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":204A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Tx_CredIVA 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   300
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta de Crédito IVA:"
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   60
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cancelación de Facturas"
         Height          =   1455
         Left            =   360
         TabIndex        =   53
         Top             =   1860
         Width           =   6915
         Begin VB.TextBox Tx_PagoFacturas 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   55
            ToolTipText     =   "Indicar medio de pago de egresos (Caja, Banco, Tarjetas, ect.)"
            Top             =   300
            Width           =   3855
         End
         Begin VB.CommandButton Bt_PagoFacturas 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":240B
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.CommandButton Bt_CobFacturas 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":27CC
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Plan de Cuentas"
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox Tx_CobFacturas 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   54
            ToolTipText     =   "Indicar medio de recaudación de dineros (Caja, Banco, Tarjetas, ect.)"
            Top             =   660
            Width           =   3855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Pago:"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "A estas cuentas se imputan, por omisión, los pagos de facturas de compra y de venta."
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   57
            Top             =   1080
            Width           =   6075
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Cobranza:"
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   56
            Top             =   720
            Width           =   1635
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Resultado Ejercicio"
         Height          =   1155
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   6915
         Begin VB.CommandButton Bt_ResEje 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":2B8D
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Plan de Cuentas"
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox Tx_ResEje 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   660
            Width           =   3795
         End
         Begin VB.TextBox Tx_Patrimonio 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   300
            Width           =   3795
         End
         Begin VB.CommandButton Bt_Patrimonio 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":2F4E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Resultado del Ejercicio:"
            Height          =   195
            Index           =   12
            Left            =   180
            TabIndex        =   52
            Top             =   720
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Patrimonio:"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Impuesto Único a los Trabajadores"
         Height          =   855
         Left            =   -74640
         TabIndex        =   45
         Top             =   4920
         Width           =   6915
         Begin VB.TextBox Tx_ImpUnico 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   300
            Width           =   3795
         End
         Begin VB.CommandButton Bt_ImpUnico 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Picture         =   "FrmConfigCtasDef.frx":330F
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Impuesto Único"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2025
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Definición de Cuentas en otros Libros"
         Height          =   3255
         Left            =   -74640
         TabIndex        =   42
         Top             =   2340
         Width           =   6915
         Begin VB.Frame Frame7 
            Caption         =   "Configuración detallada de Cuentas"
            Height          =   1455
            Left            =   240
            TabIndex        =   67
            Top             =   1500
            Width           =   6375
            Begin VB.OptionButton Op_ConfigDetLibCompras 
               Caption         =   "Cuentas por Proveedor Libro de Compras"
               Height          =   255
               Left            =   180
               TabIndex        =   70
               Top             =   420
               Value           =   -1  'True
               Width           =   3375
            End
            Begin VB.OptionButton Op_ConfigDetLibVentas 
               Caption         =   "Cuentas por Cliente Libro de Ventas"
               Height          =   255
               Left            =   180
               TabIndex        =   69
               Top             =   720
               Width           =   3375
            End
            Begin VB.CommandButton Bt_ConfigDetCtas 
               Caption         =   "Configurar..."
               Height          =   375
               Left            =   4440
               TabIndex        =   68
               Top             =   540
               Width           =   1635
            End
         End
         Begin VB.ComboBox Cb_TipoLib 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   420
            Width           =   2235
         End
         Begin VB.ComboBox Cb_TipoValor 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   900
            Width           =   2235
         End
         Begin VB.CommandButton Bt_AsignarCuentas 
            Caption         =   "Ver / Asignar Cuentas"
            Height          =   795
            Left            =   4920
            Picture         =   "FrmConfigCtasDef.frx":36D0
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Libro:"
            Height          =   255
            Left            =   180
            TabIndex        =   44
            Top             =   540
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Valor:"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   43
            Top             =   960
            Width           =   1035
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Libro de Retenciones y Honorarios"
         Height          =   1575
         Index           =   2
         Left            =   -74640
         TabIndex        =   35
         Top             =   600
         Width           =   6915
         Begin VB.CommandButton Bt_ImpRet 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":3C6B
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Tx_ImpRet 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   300
            Width           =   3795
         End
         Begin VB.TextBox Tx_NetoHon 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   660
            Width           =   3795
         End
         Begin VB.CommandButton Bt_NetoHon 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":402C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Plan de Cuentas"
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton Bt_NetoDieta 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":43ED
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Plan de Cuentas"
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox Tx_NetoDieta 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1020
            Width           =   3795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Impuesto Retenido:"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   41
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Neto Honorarios:"
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   40
            Top             =   720
            Width           =   2115
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Neto Dieta o Partic.:"
            Height          =   195
            Index           =   14
            Left            =   180
            TabIndex        =   39
            Top             =   1080
            Width           =   2355
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Otros Impuestos"
         Height          =   2415
         Index           =   1
         Left            =   -74640
         TabIndex        =   28
         Top             =   2400
         Width           =   6915
         Begin VB.Frame Frame8 
            Height          =   795
            Left            =   180
            TabIndex        =   71
            Top             =   1260
            Width           =   6555
            Begin VB.OptionButton Op_TipoLibConfig 
               Caption         =   "Ventas"
               Height          =   195
               Index           =   2
               Left            =   4020
               TabIndex        =   75
               Top             =   360
               Width           =   1035
            End
            Begin VB.OptionButton Op_TipoLibConfig 
               Caption         =   "Compras"
               Height          =   195
               Index           =   1
               Left            =   2760
               TabIndex        =   74
               Top             =   360
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.CommandButton Bt_ConfigOtrosImp 
               Caption         =   "Configurar..."
               Height          =   375
               Left            =   5100
               TabIndex        =   73
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label Label4 
               Caption         =   "Configurar Detalle Otros Impuestos"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   360
               Width           =   2595
            End
         End
         Begin VB.CommandButton Bt_OtrosImpDeb 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":47AE
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Plan de Cuentas"
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox Tx_OtrosImpDeb 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   660
            Width           =   3795
         End
         Begin VB.TextBox Tx_OtrosImpCred 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   300
            Width           =   3795
         End
         Begin VB.CommandButton Bt_OtrosImpCred 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":4B6F
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Puede ser la misma cuenta, si se manejan juntos."
            Height          =   195
            Index           =   4
            Left            =   2580
            TabIndex        =   34
            Top             =   1020
            Width           =   3465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Otros Imp. Débito:"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   33
            Top             =   720
            Width           =   2190
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para Otros Imp. Crédito:"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   32
            Top             =   360
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "Sólo aplicable a contribuyentes con derecho a Crédito Fiscal por imp. especiales  D.L. 825."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   10
            Left            =   180
            TabIndex        =   31
            Top             =   2100
            Width           =   6495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "I.V.A."
         Height          =   1695
         Index           =   0
         Left            =   -74640
         TabIndex        =   23
         Top             =   600
         Width           =   6915
         Begin VB.CommandButton Bt_IVADeb 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":4F30
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Plan de Cuentas"
            Top             =   660
            Width           =   315
         End
         Begin VB.TextBox Tx_IVADeb 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   660
            Width           =   3795
         End
         Begin VB.CommandButton Bt_IVACred 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6420
            Picture         =   "FrmConfigCtasDef.frx":52F1
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Plan de Cuentas"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox Tx_IVACred 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   300
            Width           =   3795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "o Art. 14 D desde el 01.01.2020"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   18
            Left            =   180
            TabIndex        =   76
            Top             =   1320
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Puede ser la misma cuenta, excepto contribuyentes art. 14 Ter A) vigente al 31.12.2019 "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   66
            Top             =   1080
            Width           =   6255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para IVA Dédito Fiscal:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   27
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta para IVA Crédito Fiscal:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   360
            Width           =   2205
         End
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   9060
      TabIndex        =   21
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   9060
      TabIndex        =   20
      Top             =   360
      Width           =   1035
   End
End
Attribute VB_Name = "FrmConfigCtasDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdCtaIVACred As Long
Dim lIdCtaIVADeb As Long
Dim lIdCtaOtrosImpCred As Long
Dim lIdCtaOtrosImpDeb As Long
Dim lIdCtaImpRet As Long
Dim lIdCtaNetoHon As Long
Dim lIdCtaNetoDieta As Long
Dim lIdCtaPagoFacturas As Long
Dim lIdCtaCobFacturas As Long
Dim lIdCtaResEje As Long
Dim lIdCtaPatrimonio As Long
Dim lIdCtaImpUnico As Long
Dim lIdCtaCredIVA As Long
Dim lIdCtaIVAIrrec As Long
Dim lIdCtaRet3Porc As Long
Dim lIdCta3PorcCentRem As Long
'Feña
Dim lIdCtaOdfActivo As Long
Dim lIdCtaOdfPasivo As Long
' Fin Feña

'pipe 2699582
Dim lIdCtaPpmObligatorio As Long
Dim lIdCtaPpmVoluntario As Long

Private Sub Bt_AsignarCuentas_Click()
   Dim Frm As FrmPlanCuentas
   
   Set Frm = New FrmPlanCuentas
   Call Frm.FSelCuentasBasicas(ItemData(Cb_TipoLib), ItemData(Cb_TipoValor))
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me

End Sub


Private Sub Bt_ConfigDetCtas_Click()
   Dim Frm As Form
   
   If Op_ConfigDetLibCompras <> 0 Then
      Set Frm = New FrmConfigCtasLibCompras
      Frm.Show vbModal
      Set Frm = Nothing
   Else
      Set Frm = New FrmConfigCtasLibVentas
      Frm.Show vbModal
      Set Frm = Nothing
   End If

End Sub

Private Sub Bt_ConfigOtrosImp_Click()
   Dim Frm As FrmConfigImpAdic
   
   Set Frm = New FrmConfigImpAdic
   If Op_TipoLibConfig(LIB_COMPRAS) Then
      Call Frm.FConfig(LIB_COMPRAS)
   Else
      Call Frm.FConfig(LIB_VENTAS)
   End If
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Imp3CentraRem_Click()
Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Me.Tx_Imp3CentraRem = Descrip
      lIdCta3PorcCentRem = IdCuenta
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_ImpUnico_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_ImpUnico = Descrip
      lIdCtaImpUnico = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_IVAIrrec_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_IVAIrrec = Descrip
      lIdCtaIVAIrrec = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_OdfActivo_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Me.Tx_OdfActivo = Descrip
      lIdCtaOdfActivo = IdCuenta
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_OdfPasivo_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Me.Tx_OdfPasivo = Descrip
      lIdCtaOdfPasivo = IdCuenta
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_OK_Click()
   If Not Valida() Then
      Exit Sub
   End If
   
   Call SaveAll
   
   Unload Me
 
End Sub

Private Sub Bt_PagoFacturas_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_PagoFacturas = Descrip
      lIdCtaPagoFacturas = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_CobFacturas_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_CobFacturas = Descrip
      lIdCtaCobFacturas = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_CredIVA_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_CredIVA = Descrip
      lIdCtaCredIVA = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_PpmOblig_Click()
 Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Me.Tx_PpmObligatorio = Descrip
      lIdCtaPpmObligatorio = IdCuenta
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_PpmVolunt_Click()
Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Me.Tx_PpmVoluntario = Descrip
      lIdCtaPpmVoluntario = IdCuenta
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Bt_ResEje_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_ResEje = Descrip
      lIdCtaResEje = IdCuenta
   End If
   
   Set Frm = Nothing


End Sub
Private Sub Bt_Patrimonio_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   Dim ClasCta As Integer
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre, False) = vbOK Then
      ClasCta = GetClasCuenta(IdCuenta)
      
      If ClasCta <> CLASCTA_PASIVO Then
         MsgBox1 "La cuenta de Patrimonio debe ser una cuenta de Pasivo.", vbExclamation
         Exit Sub
      End If
      
      Tx_Patrimonio = Descrip
      lIdCtaPatrimonio = IdCuenta
   End If
   
   Set Frm = Nothing


End Sub

Private Sub Bt_Ret3Porc_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_Ret3Porc = Descrip
      lIdCtaRet3Porc = IdCuenta
   End If
   
   Set Frm = Nothing


End Sub

Private Sub Cb_TipoLib_Click()
   Dim i As Integer
   Dim TipoLib As Integer

   If Cb_TipoLib.ListIndex < 0 Then
      Exit Sub
   End If
   
   TipoLib = ItemData(Cb_TipoLib)
   
   Call FillTipoValLib(Cb_TipoValor, TipoLib, True, False, "CTASDEF")
   
End Sub



Private Sub Form_Load()
      
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   St_tab.Tab = 0
   
   Call SetTxRO(Tx_IVACred, True)
   Call SetTxRO(Tx_IVADeb, True)
   Call SetTxRO(Tx_OtrosImpCred, True)
   Call SetTxRO(Tx_OtrosImpDeb, True)
   Call SetTxRO(Tx_ImpRet, True)
   Call SetTxRO(Tx_NetoHon, True)
   Call SetTxRO(Tx_NetoDieta, True)
   Call SetTxRO(Tx_Patrimonio, True)
   Call SetTxRO(Tx_ResEje, True)
   Call SetTxRO(Tx_PagoFacturas, True)
   Call SetTxRO(Tx_CobFacturas, True)
   Call SetTxRO(Tx_CredIVA, True)
   Call SetTxRO(Tx_ImpUnico, True)
   Call SetTxRO(Tx_IVAIrrec, True)
   Call SetTxRO(Tx_Ret3Porc, True)
   Call SetTxRO(Tx_OdfActivo, True)
   Call SetTxRO(Tx_OdfPasivo, True)
   Call SetTxRO(Me.Tx_Imp3CentraRem, True)
   
   'pipe 2699582
   
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
   Fr_ppm.visible = True
   
   
   Call SetTxRO(Tx_PpmObligatorio, True)
   Call SetTxRO(Tx_PpmVoluntario, True)
   Else
   
   Fr_ppm.visible = False
   
   End If
   'fin 2699582
   
   Cb_TipoLib.AddItem gTipoLib(LIB_COMPRAS)
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_COMPRAS
   Cb_TipoLib.AddItem gTipoLib(LIB_VENTAS)
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_VENTAS
   Cb_TipoLib.AddItem gTipoLib(LIB_RETEN)
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_RETEN
   Cb_TipoLib.ListIndex = 0
   
   Fr_IVAIrrec.visible = gFunciones.ProporcionalidadIVA
   
   Fr_Ret3Porc.visible = IIf(gEmpresa.Ano >= 2021 And gEmpresa.Ano <= 2024, True, False)
   
   Call LoadAll
   
   Call SetupPriv
   
End Sub

Private Sub LoadAll()

   If gCtasBas.IdCtaIVACred > 0 Then
      Tx_IVACred = GetDescCuenta(gCtasBas.IdCtaIVACred)
      lIdCtaIVACred = gCtasBas.IdCtaIVACred
   End If
   
   If gCtasBas.IdCtaIVADeb > 0 Then
      Tx_IVADeb = GetDescCuenta(gCtasBas.IdCtaIVADeb)
      lIdCtaIVADeb = gCtasBas.IdCtaIVADeb
   End If
   
   If gCtasBas.IdCtaOtrosImpCred > 0 Then
      Tx_OtrosImpCred = GetDescCuenta(gCtasBas.IdCtaOtrosImpCred)
      lIdCtaOtrosImpCred = gCtasBas.IdCtaOtrosImpCred
   End If
  
   If gCtasBas.IdCtaOtrosImpDeb > 0 Then
      Tx_OtrosImpDeb = GetDescCuenta(gCtasBas.IdCtaOtrosImpDeb)
      lIdCtaOtrosImpDeb = gCtasBas.IdCtaOtrosImpDeb
   End If
   
   If gCtasBas.IdCtaImpRet > 0 Then
      Tx_ImpRet = GetDescCuenta(gCtasBas.IdCtaImpRet)
      lIdCtaImpRet = gCtasBas.IdCtaImpRet
   End If
   
   If gCtasBas.IdCtaNetoHon > 0 Then
      Tx_NetoHon = GetDescCuenta(gCtasBas.IdCtaNetoHon)
      lIdCtaNetoHon = gCtasBas.IdCtaNetoHon
   End If
   
   If gCtasBas.IdCtaImpUnico > 0 Then
      Tx_ImpUnico = GetDescCuenta(gCtasBas.IdCtaImpUnico)
      lIdCtaImpUnico = gCtasBas.IdCtaImpUnico
   End If
   
   If gCtasBas.IdCtaNetoDieta > 0 Then
      Tx_NetoDieta = GetDescCuenta(gCtasBas.IdCtaNetoDieta)
      lIdCtaNetoDieta = gCtasBas.IdCtaNetoDieta
   End If
   
   If gCtasBas.IdCtaPatrimonio > 0 Then
      Tx_Patrimonio = GetDescCuenta(gCtasBas.IdCtaPatrimonio)
      lIdCtaPatrimonio = gCtasBas.IdCtaPatrimonio
   End If
   
   If gCtasBas.IdCtaResEje > 0 Then
      Tx_ResEje = GetDescCuenta(gCtasBas.IdCtaResEje)
      lIdCtaResEje = gCtasBas.IdCtaResEje
   End If
   
   If gCtasBas.IdCtaPagoFacturas > 0 Then
      Tx_PagoFacturas = GetDescCuenta(gCtasBas.IdCtaPagoFacturas)
      lIdCtaPagoFacturas = gCtasBas.IdCtaPagoFacturas
   End If
  
   If gCtasBas.IdCtaCobFacturas > 0 Then
      Tx_CobFacturas = GetDescCuenta(gCtasBas.IdCtaCobFacturas)
      lIdCtaCobFacturas = gCtasBas.IdCtaCobFacturas
   End If
  
   If gCtasBas.IdCtaCredIVA > 0 Then
      Tx_CredIVA = GetDescCuenta(gCtasBas.IdCtaCredIVA)
      lIdCtaCredIVA = gCtasBas.IdCtaCredIVA
   End If
   
   If gCtasBas.IdCtaIVAIrrec > 0 Then
      Tx_IVAIrrec = GetDescCuenta(gCtasBas.IdCtaIVAIrrec)
      lIdCtaIVAIrrec = gCtasBas.IdCtaIVAIrrec
   End If
  
   If gCtasBas.IdCtaRet3Porc > 0 Then
      Tx_Ret3Porc = GetDescCuenta(gCtasBas.IdCtaRet3Porc)
      lIdCtaRet3Porc = gCtasBas.IdCtaRet3Porc
   End If
   
   If gCtasBas.IdCta3PorcCentraRem > 0 Then
      Me.Tx_Imp3CentraRem = GetDescCuenta(gCtasBas.IdCta3PorcCentraRem)
      lIdCta3PorcCentRem = gCtasBas.IdCta3PorcCentraRem
   End If
   
   'feña
   If gCtasBas.IdCtaOdfActivo > 0 Then
      Me.Tx_OdfActivo = GetDescCuenta(gCtasBas.IdCtaOdfActivo)
      lIdCtaOdfActivo = gCtasBas.IdCtaOdfActivo
   End If
   
   If gCtasBas.IdCtaOdfPasivo > 0 Then
      Me.Tx_OdfPasivo = GetDescCuenta(gCtasBas.IdCtaOdfPasivo)
      lIdCtaOdfPasivo = gCtasBas.IdCtaOdfPasivo
   End If
   'fin feña
  
   'pipe 2699582
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
   
    If gCtasBas.IdCtaPpmObligatorio > 0 Then
      Me.Tx_PpmObligatorio = GetDescCuenta(gCtasBas.IdCtaPpmObligatorio)
      lIdCtaPpmObligatorio = gCtasBas.IdCtaPpmObligatorio
   End If


    If gCtasBas.IdCtaPpmVoluntario > 0 Then
      Me.Tx_PpmVoluntario = GetDescCuenta(gCtasBas.IdCtaPpmVoluntario)
      lIdCtaPpmVoluntario = gCtasBas.IdCtaPpmVoluntario
   End If
   
   End If
   'fin 2699582

End Sub
Private Sub Bt_IVACred_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_IVACred = Descrip
      lIdCtaIVACred = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_IVADeb_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_IVADeb = Descrip
      lIdCtaIVADeb = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_OtrosImpCred_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_OtrosImpCred = Descrip
      lIdCtaOtrosImpCred = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_OtrosImpDeb_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_OtrosImpDeb = Descrip
      lIdCtaOtrosImpDeb = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_ImpRet_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_ImpRet = Descrip
      lIdCtaImpRet = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub

Private Sub Bt_NetoHon_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_NetoHon = Descrip
      lIdCtaNetoHon = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub
Private Sub Bt_NetoDieta_Click()
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre) = vbOK Then
      Tx_NetoDieta = Descrip
      lIdCtaNetoDieta = IdCuenta
   End If
   
   Set Frm = Nothing

End Sub
Private Sub SaveAll()
   Dim Rc As Long
   Dim Rs As Recordset
         
   'cuentas IVA y Otros Impuestos
         
   'actualizamos
   If gCtasBas.IdCtaIVACred <> lIdCtaIVACred Then
      Call UpdParamEmpresa("CTAIVACRED", 0, lIdCtaIVACred)
   End If
   
   If gCtasBas.IdCtaIVADeb <> lIdCtaIVADeb Then
      Call UpdParamEmpresa("CTAIVADEB", 0, lIdCtaIVADeb)
   End If
   
   If gCtasBas.IdCtaOtrosImpCred <> lIdCtaOtrosImpCred Then
      Call UpdParamEmpresa("CTAOIMPCRE", 0, lIdCtaOtrosImpCred)
   End If
   
   If gCtasBas.IdCtaOtrosImpDeb <> lIdCtaOtrosImpDeb Then
      Call UpdParamEmpresa("CTAOIMPDEB", 0, lIdCtaOtrosImpDeb)
   End If
   
   If gCtasBas.IdCtaPagoFacturas <> lIdCtaPagoFacturas Then
      Call UpdParamEmpresa("CTAPAGOFAC", 0, lIdCtaPagoFacturas)
   End If
   
   If gCtasBas.IdCtaCobFacturas <> lIdCtaCobFacturas Then
      Call UpdParamEmpresa("CTACOBFAC", 0, lIdCtaCobFacturas)
   End If
   
   If gCtasBas.IdCtaCredIVA <> lIdCtaCredIVA Then
      Call UpdParamEmpresa("CTACREDIVA", 0, lIdCtaCredIVA)
   End If
   
   If gCtasBas.IdCtaIVAIrrec <> lIdCtaIVAIrrec Then
      Call UpdParamEmpresa("CTAIVAIRRE", 0, lIdCtaIVAIrrec)
   End If

   If gCtasBas.IdCtaImpRet <> lIdCtaImpRet Then
      Call UpdParamEmpresa("CTAIMPRET", 0, lIdCtaImpRet)
   End If
   
   If gCtasBas.IdCtaNetoHon <> lIdCtaNetoHon Then
      Call UpdParamEmpresa("CTANETORET", 0, lIdCtaNetoHon)
   End If
   
   If gCtasBas.IdCtaNetoDieta <> lIdCtaNetoDieta Then
      Call UpdParamEmpresa("CTANETODIE", 0, lIdCtaNetoDieta)
   End If
            
   If gCtasBas.IdCtaPatrimonio <> lIdCtaPatrimonio Then
      Call UpdParamEmpresa("CTAPATRIM", 0, lIdCtaPatrimonio)
   End If
   
   If gCtasBas.IdCtaResEje <> lIdCtaResEje Then
      Call UpdParamEmpresa("CTARESEJE", 0, lIdCtaResEje)
   End If
            
   If gCtasBas.IdCtaRet3Porc <> lIdCtaRet3Porc Then
      Call UpdParamEmpresa("CTARET3PRC", 0, lIdCtaRet3Porc)
   End If
   
   If gCtasBas.IdCta3PorcCentraRem <> lIdCta3PorcCentRem Then
      Call UpdParamEmpresa("CTA3CENREM", 0, lIdCta3PorcCentRem)
   End If
   
   If gCtasBas.IdCtaOdfActivo <> lIdCtaOdfActivo Then
      Call UpdParamEmpresa("CTAODFACTI", 0, lIdCtaOdfActivo)
   End If
   
    If gCtasBas.IdCtaOdfPasivo <> lIdCtaOdfPasivo Then
      Call UpdParamEmpresa("CTAODFPASI", 0, lIdCtaOdfPasivo)
   End If
            
   If gCtasBas.IdCtaImpUnico <> lIdCtaImpUnico Then
      Call UpdParamEmpresa("CTAIMPUNIC", 0, lIdCtaImpUnico)
            
      If gCtasBas.IdCtaImpUnico > 0 Then
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & lIdCtaImpUnico & "' WHERE Tipo = 'CTAIMPUNIC' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
 
      Else
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES ('CTAIMPUNIC', 0, '" & lIdCtaImpUnico & "', " & gEmpresa.id & "," & gEmpresa.Ano & ")")
      End If
      
      'actuliazamos el código F29 del plan de cuentas
      'limpiamos el anterior si corresponde
      If gCtasBas.IdCtaImpUnico > 0 Then
         Call ExecSQL(DbMain, "UPDATE Cuentas SET CodF29=0 WHERE IdCuenta = " & gCtasBas.IdCtaImpUnico & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      End If
      Call ExecSQL(DbMain, "UPDATE Cuentas SET CodF29=" & CODF29_IMPUNICO & " WHERE IdCuenta = " & lIdCtaImpUnico & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   End If
   
   'pipe 2699582
   
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
       If gCtasBas.IdCtaPpmObligatorio <> lIdCtaPpmObligatorio Then
          Call UpdParamEmpresa("CTAPPMOBLI", 0, lIdCtaPpmObligatorio)
       End If
    
        If gCtasBas.IdCtaPpmVoluntario <> lIdCtaPpmVoluntario Then
          Call UpdParamEmpresa("CTAPPMVOLU", 0, lIdCtaPpmVoluntario)
       End If
            
    End If
   'fin 2699582
            
   gCtasBas.IdCtaIVACred = lIdCtaIVACred
   gCtasBas.IdCtaIVADeb = lIdCtaIVADeb
   gCtasBas.IdCtaOtrosImpCred = lIdCtaOtrosImpCred
   gCtasBas.IdCtaOtrosImpDeb = lIdCtaOtrosImpDeb
   gCtasBas.IdCtaPagoFacturas = lIdCtaPagoFacturas
   gCtasBas.IdCtaCobFacturas = lIdCtaCobFacturas
   gCtasBas.IdCtaImpRet = lIdCtaImpRet
   gCtasBas.IdCtaNetoHon = lIdCtaNetoHon
   gCtasBas.IdCtaNetoDieta = lIdCtaNetoDieta
   gCtasBas.IdCtaPatrimonio = lIdCtaPatrimonio
   gCtasBas.IdCtaResEje = lIdCtaResEje
   gCtasBas.IdCtaImpUnico = lIdCtaImpUnico
   gCtasBas.IdCtaCredIVA = lIdCtaCredIVA
   gCtasBas.IdCtaIVAIrrec = lIdCtaIVAIrrec
   gCtasBas.IdCtaRet3Porc = lIdCtaRet3Porc
   gCtasBas.IdCta3PorcCentraRem = lIdCta3PorcCentRem
   gCtasBas.IdCtaOdfActivo = lIdCtaOdfActivo
   gCtasBas.IdCtaOdfPasivo = lIdCtaOdfPasivo
   
   'pipe 2699582
   
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
   gCtasBas.IdCtaPpmObligatorio = lIdCtaPpmObligatorio
   gCtasBas.IdCtaPpmVoluntario = lIdCtaPpmVoluntario
   
   End If
   'fin 2699582
   
End Sub
Private Function Valida() As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   
   Valida = False
      
   'Cuentas IVA y Otros Impuestos
   If (lIdCtaIVACred <= 0) Then
      MsgBox1 "Falta definir la cuenta de IVA Crédito Fiscal.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaIVADeb <= 0) Then
      MsgBox1 "Falta definir la cuenta de IVA Débito Fiscal.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   'pipe 2699582
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
        If lIdCtaPpmObligatorio > 0 And lIdCtaPpmObligatorio = lIdCtaPpmVoluntario Then
            MsgBox1 "Cuenta Ppm Obligatoria no puede ser iguala Ppm Voluntario.", vbInformation + vbOKOnly
            lIdCtaPpmVoluntario = 0
            Me.Tx_PpmVoluntario = ""
            Exit Function
        End If
    
        If lIdCtaPpmVoluntario > 0 And lIdCtaPpmVoluntario = lIdCtaPpmObligatorio Then
            MsgBox1 "Cuenta Ppm Voluntario no puede ser iguala Ppm Obligatoria.", vbInformation + vbOKOnly
             lIdCtaPpmVoluntario = 0
              Me.Tx_PpmVoluntario = ""
            Exit Function
        End If
    End If
    ' fin pipe 2699582

   
    If gEmpresa.Franq14Ter Then
      If lIdCtaIVACred = lIdCtaIVADeb Then
         MsgBox1 " Contribuyentes de 14 TER debe usar cuentas contables distintas para el control del IVA Crédito e IVA Débito Fiscal.", vbExclamation + vbOKOnly
         Exit Function
      End If
   End If
   
  
   
   
'   If (lIdCtaOtrosImpCred <= 0) Then
'      MsgBox1 "Falta definir la cuenta de Otros Impuestos Crédito.", vbExclamation + vbOKOnly
'      Exit Function
'   End If
'
'   If (lIdCtaOtrosImpDeb <= 0) Then
'      MsgBox1 "Falta definir la cuenta de Otros Impuestos Dédito.", vbExclamation + vbOKOnly
'      Exit Function
'   End If
   
   If (lIdCtaImpRet <= 0) Then
      MsgBox1 "Falta definir la cuenta de Impuesto Retenido.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaNetoHon <= 0) Then
      MsgBox1 "Falta definir la cuenta de Neto Honorarios.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaNetoHon <= 0) Then
      MsgBox1 "Falta definir la cuenta de Neto Dieta o Participación.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaPatrimonio <= 0) Then
      MsgBox1 "Falta definir la cuenta de Patrimonio.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaResEje <= 0) Then
      MsgBox1 "Falta definir la cuenta de Resultado Ejercicio.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaPagoFacturas <= 0) Then
      MsgBox1 "Falta definir la cuenta de Pago Facturas.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaCobFacturas <= 0) Then
      MsgBox1 "Falta definir la cuenta de Cobranza Facturas.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaImpUnico <= 0) Then
      MsgBox1 "Falta definir la cuenta de Impuesto Único a los Trabajadores.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If (lIdCtaCredIVA <= 0) Then
      MsgBox1 "Falta definir la cuenta de Crédito IVA para Remanente IVA año anterior.", vbExclamation + vbOKOnly
      Exit Function
   End If


   If (gCtasBas.IdCtaIVACred > 0 Or gCtasBas.IdCtaIVADeb > 0 Or gCtasBas.IdCtaOtrosImpCred > 0 Or gCtasBas.IdCtaOtrosImpDeb > 0) And (gCtasBas.IdCtaIVACred <> lIdCtaIVACred Or gCtasBas.IdCtaIVADeb <> lIdCtaIVADeb Or gCtasBas.IdCtaOtrosImpCred <> lIdCtaOtrosImpCred Or gCtasBas.IdCtaOtrosImpDeb <> lIdCtaOtrosImpDeb) Then
      
      'ya existe una definición y es distinta a la que había
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then    'hay al menos un documento
      
         'MsgBox1 "No es posible cambiar las cuentas asignadas a IVA y Otros Impuestos, hay documentos ya ingresados.", vbExclamation + vbOKOnly
         If MsgBox1("ATENCIÓN:" & vbNewLine & vbNewLine & "Si cambia las cuentas asignadas a IVA y Otros Impuestos, no se actualizarán estas cuentas en los documentos ya ingresados." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Call CloseRs(Rs)
            Exit Function
         End If
      
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If (gCtasBas.IdCtaImpRet > 0 Or gCtasBas.IdCtaNetoHon > 0 Or gCtasBas.IdCtaNetoDieta > 0) And (gCtasBas.IdCtaImpRet <> lIdCtaImpRet Or gCtasBas.IdCtaNetoHon <> lIdCtaNetoHon Or gCtasBas.IdCtaNetoDieta <> lIdCtaNetoDieta) Then
      'ya existe una definición y es distinta a la que había
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib =" & LIB_RETEN
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then    'hay al menos un documento
      
         'MsgBox1 "No es posible cambiar las cuentas asignadas a Retenciones y Honorarios, hay documentos ya ingresados.", vbExclamation + vbOKOnly
         If MsgBox1("ATENCIÓN:" & vbNewLine & vbNewLine & "Si cambia las cuentas asignadas a Retenciones y Honorarios, no se actualizarán estas cuentas en los documentos ya ingresados." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Call CloseRs(Rs)
            Exit Function
         End If
      
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   
        If lIdCtaOdfActivo > 0 And lIdCtaOdfActivo = lIdCtaOdfPasivo Then
            MsgBox1 "Cuenta ODF Activo no puede ser igual a cuenta ODF Pasivo.", vbInformation + vbOKOnly
            lIdCtaOdfActivo = 0
            Me.Tx_OdfActivo = ""
            Exit Function
        End If
    
        If lIdCtaOdfPasivo > 0 And lIdCtaOdfPasivo = lIdCtaOdfActivo Then
            MsgBox1 "Cuenta ODF Pasivo no puede ser iguala cuenta ODF Activo.", vbInformation + vbOKOnly
             lIdCtaOdfPasivo = 0
              Me.Tx_OdfPasivo = ""
            Exit Function
        End If
    
   
   Valida = True
End Function

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   
End Function



