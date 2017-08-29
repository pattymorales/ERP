VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{DECFDD03-1E4B-11D1-B65E-0000C039C248}#5.0#0"; "mskedit.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FOrdenVenta 
   Caption         =   "Subtotal:"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   4800
      Width           =   7410
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   30
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmCabecera 
      Caption         =   "Cabecera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3520
         MaxLength       =   64
         TabIndex        =   14
         Top             =   540
         Width           =   4000
      End
      Begin VB.TextBox txtorden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   64
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtCcliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   12
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtdireccion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   11
         Top             =   840
         Width           =   5715
      End
      Begin VB.TextBox txttelefono 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1140
         Width           =   1275
      End
      Begin MskeditLib.MaskInBox mskfecha 
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   1440
         Width           =   975
         _Version        =   262144
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "__/__/____"
         Appearance      =   1
         Decimals        =   2
         Separator       =   -1  'True
         MaskType        =   2
         HideSelection   =   0   'False
         MaxLength       =   0
         AutoTab         =   0   'False
         DateString      =   "__/__/____"
         FormattedText   =   ""
         Mask            =   "##/##/####"
         HelpLine        =   ""
         ClipText        =   ""
         ClipMode        =   0
         StringIndex     =   0
         DateType        =   0
         DateSybase      =   "03/24/11"
         AutoDecimal     =   0   'False
         MinReal         =   -1.1e38
         MaxReal         =   3.4e38
         Units           =   0
         Errores         =   0
      End
      Begin Threed.SSCommand cmdBuscarCliente 
         Height          =   285
         Left            =   3090
         TabIndex        =   21
         Tag             =   "21315"
         ToolTipText     =   "Buscar"
         Top             =   540
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":25C2
      End
      Begin VB.Label lblnumero 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número de Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código del Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Venta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblTelefono 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   675
      End
   End
   Begin VB.Frame frmlinea 
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   7695
      Begin VB.TextBox txtCarticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtAdescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2800
         MaxLength       =   64
         TabIndex        =   5
         Top             =   240
         Width           =   4380
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   4
         Top             =   540
         Width           =   1275
      End
      Begin Threed.SSCommand cmdRuteo 
         Height          =   360
         Index           =   1
         Left            =   7200
         TabIndex        =   23
         Top             =   240
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":2E9C
      End
      Begin Threed.SSCommand cmdRuteo 
         Height          =   360
         Index           =   2
         Left            =   7200
         TabIndex        =   24
         Top             =   615
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":3776
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   285
         Left            =   2370
         TabIndex        =   25
         Tag             =   "21315"
         ToolTipText     =   "Buscar"
         Top             =   240
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FOrdenVenta.frx":4050
      End
      Begin VB.Label lblArticulo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblCantidad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   675
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6550
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3405
      Width           =   1275
   End
   Begin MSGrid.Grid Grid1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   2355
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdImprimir 
      Height          =   600
      Left            =   7920
      TabIndex        =   22
      ToolTipText     =   "Imprimir"
      Top             =   2040
      Width           =   900
      _Version        =   65536
      _ExtentX        =   1587
      _ExtentY        =   1058
      _StockProps     =   78
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FOrdenVenta.frx":492A
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5760
      TabIndex        =   0
      Top             =   3405
      Width           =   405
   End
   Begin VB.Line linLinea 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   0
      X2              =   11040
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "FOrdenVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscarCliente_Click()
    FBuscarCliente.Show 1
End Sub


Private Sub jmbSalir_Click()
    Unload Me
End Sub
