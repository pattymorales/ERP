VERSION 5.00
Object = "{DECFDD03-1E4B-11D1-B65E-0000C039C248}#5.0#0"; "mskedit.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FAutorizaciones 
   Caption         =   "Autorizaciones"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   7410
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   10
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
         Picture         =   "FAutorizaciones.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   11
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
         Picture         =   "FAutorizaciones.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   12
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
         Picture         =   "FAutorizaciones.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   13
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
         Picture         =   "FAutorizaciones.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
         TabIndex        =   14
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
         Picture         =   "FAutorizaciones.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos Generales"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.ComboBox cmbEstatus 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtcodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   64
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin MskeditLib.MaskInBox mskfechavalidez 
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   600
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
      Begin MskeditLib.MaskInBox mskfechaautoriza 
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   900
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
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Estatus de la Autorización:"
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
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblFechaautoriza 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de la Autorización:"
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
         Top             =   900
         Width           =   2190
      End
      Begin VB.Label lblFechavalidez 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Validez:"
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
         TabIndex        =   6
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblNumero 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número de Autorización SRI:"
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
         TabIndex        =   5
         Top             =   300
         Width           =   2475
      End
   End
   Begin JMButton.JMBcontrol jmbDetalle 
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "FAutorizaciones.frx":25C2
      Caption         =   "Detalles"
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub jmbDetalle_Click()
    FLineaautorizacion.Show 0
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub
