VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.MDIForm FPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "ERP"
   ClientHeight    =   6480
   ClientLeft      =   240
   ClientTop       =   810
   ClientWidth     =   9450
   HelpContextID   =   1
   Icon            =   "Fprincip.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin Threed.SSPanel pnlBarraAyuda 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9450
      _Version        =   65536
      _ExtentX        =   16669
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Begin VB.Timer trmMail 
         Enabled         =   0   'False
         Left            =   4635
         Top             =   15
      End
      Begin VB.Timer tmrHora 
         Interval        =   60000
         Left            =   5850
         Top             =   -30
      End
   End
   Begin Threed.SSPanel pnlBarraMensajes 
      Align           =   2  'Align Bottom
      Height          =   348
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   9444
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Begin Threed.SSPanel pnlTransaccionLine 
         Height          =   255
         Left            =   4230
         TabIndex        =   9
         Top             =   45
         Width           =   3660
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel Focos 
         Height          =   135
         Index           =   3
         Left            =   9180
         TabIndex        =   7
         Top             =   165
         Width           =   195
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "P"
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Outline         =   -1  'True
         FloodType       =   1
         FloodColor      =   65280
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel Focos 
         Height          =   135
         Index           =   2
         Left            =   8970
         TabIndex        =   6
         Top             =   165
         Width           =   195
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "R"
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Outline         =   -1  'True
         FloodType       =   1
         FloodColor      =   65280
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel Focos 
         Height          =   135
         Index           =   1
         Left            =   8760
         TabIndex        =   5
         Top             =   165
         Width           =   195
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "T"
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Outline         =   -1  'True
         FloodType       =   1
         FloodColor      =   65280
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel Focos 
         Height          =   135
         Index           =   0
         Left            =   8550
         TabIndex        =   4
         Top             =   165
         Width           =   195
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "L"
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Outline         =   -1  'True
         FloodType       =   1
         FloodColor      =   65280
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel pnlHora 
         Height          =   250
         Left            =   7905
         TabIndex        =   2
         Top             =   45
         Width           =   555
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlHelpLine 
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   45
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   1
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   " L   T   R  P"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   8535
         TabIndex        =   8
         Top             =   30
         Width           =   870
      End
   End
   Begin VB.Menu mnuConexion 
      Caption         =   "Cone&xión"
      Begin VB.Menu mnuLogon 
         Caption         =   "Log o&n"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log o&ff"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLinea0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuPasswd 
         Caption         =   "Password"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuArticulos 
      Caption         =   "&Artículos"
      Begin VB.Menu mnuDatosga 
         Caption         =   "&Datos Generales de Artículos"
         Begin VB.Menu mnuFamilia 
            Caption         =   "&Famila de Artículos"
         End
         Begin VB.Menu mnuUnidades 
            Caption         =   "&Unidades de Stock"
         End
      End
      Begin VB.Menu mnuLinea4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManarticulos 
         Caption         =   "&Mantener Artículos"
      End
      Begin VB.Menu mnuLinea5 
         Caption         =   "-"
      End
      Begin VB.Menu mnucorrecion 
         Caption         =   "&Corrección de Stock"
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuOrdenVenta 
         Caption         =   "Orden de Venta"
      End
   End
   Begin VB.Menu mnuDatos 
      Caption         =   "&Datos Generales"
      Begin VB.Menu mnuCompanias 
         Caption         =   "&Compañías"
      End
      Begin VB.Menu mnuSucursal 
         Caption         =   "&Sucursales"
      End
      Begin VB.Menu mnuLinea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnudatoscliente 
         Caption         =   "&Datos Clientes"
      End
   End
   Begin VB.Menu mnuFacturacion 
      Caption         =   "&Facturación"
      Begin VB.Menu mnuAutorizacion 
         Caption         =   "Autorizaciones"
      End
   End
   Begin VB.Menu mnuVentanas 
      Caption         =   "&Ventanas"
      WindowList      =   -1  'True
      Begin VB.Menu mnuBarraAyuda 
         Caption         =   "Barra de ayuda"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBarraMensajes 
         Caption         =   "Barra de mensajes"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLinea9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascada 
         Caption         =   "Cascada"
      End
      Begin VB.Menu mnuHorizontal 
         Caption         =   "Horizontal"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Vertical"
      End
      Begin VB.Menu mnuIconos 
         Caption         =   "Organizar Iconos"
      End
      Begin VB.Menu mnuLinea10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSizetoFit 
         Caption         =   "Ajustar la ventana"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "Contenido"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuLinea11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "FPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DescargarFormas()
'RAN: 15-abr-1999
'Cambio del código de descarga de las formas de forma
'manual a forma automatica movimiendonos por las instancias
'de la la Forms

    Dim VTForm As Form
    For Each VTForm In Forms
        Unload VTForm
    Next
End Sub

Private Sub MDIForm_Load()
    
    FPrincipal!Focos(0).FloodPercent = 100
    FPrincipal!Focos(1).FloodPercent = 0
    FPrincipal!Focos(2).FloodPercent = 0
    FPrincipal!Focos(3).FloodPercent = 0
    
    FPrincipal.Left = 15
    FPrincipal.Top = 15
    FPrincipal.Width = 9570
    FPrincipal.Height = 7170
    
    ' Inicializacion de las variables
    ARCHIVOINI$ = App.Path + "\segurid.ini"
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    mnuSalir_Click
End Sub

Private Sub mnuAcerca_Click()
    FAbout.Show 1
End Sub

Private Sub mnuActHor_Click()
    Operacion% = ACTUALIZAR%
    FGTiposHorario.cmdBoton(2).Visible = True
    FGTiposHorario.cmdBoton(3).Visible = False
    FGTiposHorario.HelpContextID = 410
    FGTiposHorario.Show 1
End Sub

Private Sub mnuActRol_Click()
    Operacion% = ACTUALIZAR%
    FGRoles.cmdBoton(2).Visible = True
    FGRoles.cmdBoton(3).Visible = False
    FGRoles.HelpContextID = 310
    FGRoles.Show 1
End Sub

Private Sub mnuAsiCar_Click()
    FAsiCargos.Show
End Sub

Private Sub mnuAutorizacion_Click()
    FAutorizaciones.Show 0
End Sub

Private Sub mnuBarraAyuda_Click()
    If mnuBarraAyuda.Checked = False Then
        pnlBarraAyuda.Visible = True
        mnuBarraAyuda.Checked = True
    Else
        pnlBarraAyuda.Visible = False
        mnuBarraAyuda.Checked = False
    End If
End Sub

Private Sub mnuBarraMensajes_Click()
    If mnuBarraMensajes.Checked = False Then
        pnlBarraMensajes.Visible = True
        mnuBarraMensajes.Checked = True
    Else
        pnlBarraMensajes.Visible = False
        mnuBarraMensajes.Checked = False
    End If
End Sub

Private Sub mnuBloquear_Click()
'----VQU:05-10-2010-----
'Modificaciones para NCS
'-----------------------
'    Main
'    FPrincipal.Show
'    FPrincipal.WindowState = 0

    Me.Hide
    fpantalla.Show
    
End Sub

Private Sub mnuCalFer_Click()
    FCalendario.Show
End Sub

Private Sub mnuCascada_Click()
    FPrincipal.Arrange 0
End Sub

Private Sub mnuCompanias_Click()
    FCompania.Show 0
End Sub

Private Sub mnuContenido_Click()
' CRA28abr9802 UI: Dimensionar la variable VT
Dim VT
' CRA28abr9802 UF
    VT = Shell("winhelp " + App.Path + "\segurid.hlp", 1)
    
End Sub

Private Sub mnuCopiaRol_Click()
    FCopiaroles.Show
End Sub

Private Sub mnuCProRol_Click()
    FConsProductosRol.Show
End Sub

Private Sub mnuCTraAut_Click()
    FConsTranAut.Show
End Sub

Private Sub mnuDistrib_Click()
    FDistribucion.Show
End Sub

Private Sub mnucorrecion_Click()
    FCStock.Show 0
End Sub

Private Sub mnudatoscliente_Click()
    FApCliente.Show 0
End Sub

Private Sub mnuDirecciones_Click()
    FApGeneral.Show 0
End Sub

Private Sub mnuFamilia_Click()
    FFamilia.Show 0
End Sub

Private Sub mnuHorizontal_Click()
    FPrincipal.Arrange 1
End Sub

Private Sub mnuIconos_Click()
    FPrincipal.Arrange 3
End Sub

Private Sub mnuLogin_Click()
    FUsuarioLogin.Show 1
End Sub

Private Sub mnuLogout_Click()
   FLogout.Show 0
End Sub

Private Sub mnuLogoff_Click()
    VLOK = enablemenus(False, "")
End Sub

Private Sub mnuLogon_Click()
    Flogin.Show 0
End Sub

Private Sub mnuManarticulos_Click()
    FArticulo.Show 0
End Sub

Private Sub mnuOrdenVenta_Click()
    FOrdenVenta.Show 0
End Sub

Private Sub mnuPasswd_Click()
    FCambioPassword.Show 0
End Sub

Private Sub mnuPreferencias_Click()
    FPreferencias.Show 1
End Sub


Private Sub mnuSalir_Click()
    End
End Sub

Private Sub mnuSizetoFit_Click()
    If Screen.ActiveForm.WindowState = 2 Then
        Exit Sub
    End If

    If Screen.ActiveForm.Caption = "C.O.B.I.S. - Subsistema de Seguridad" Then
        FPrincipal.Left = 15
        FPrincipal.Top = 15
        FPrincipal.Width = 9570
        FPrincipal.Height = 7170
    Else
        If Screen.ActiveForm.WindowState = 0 Then
            If Screen.ActiveForm.Caption <> "Calendario" Then
                Screen.ActiveForm.Left = 15
                Screen.ActiveForm.Top = 15
                Screen.ActiveForm.Width = 9420
                Screen.ActiveForm.Height = 5730
            Else
                Screen.ActiveForm.Left = 3000
                Screen.ActiveForm.Top = 15
                Screen.ActiveForm.Width = 4080
                Screen.ActiveForm.Height = 5730
            End If
        Else
            FPrincipal.Left = 15
            FPrincipal.Top = 15
            FPrincipal.Width = 9570
            FPrincipal.Height = 7170
        End If
    End If
End Sub

Private Sub mnuUsuNod_Click()
    FConsUsuNodos.Show
End Sub

Private Sub mnuUsuRol_Click()
    FConsUsuRol.Show
End Sub

Private Sub mnuSucursal_Click()
    FSucursal.Show 0
End Sub

Private Sub mnuUnidades_Click()
    FUnidades.Show 0
End Sub

Private Sub mnuUsuarios_Click()
    Fusuario.Show 0
End Sub

Private Sub mnuVertical_Click()
    FPrincipal.Arrange 2
End Sub

Private Sub tmrHora_Timer()
    pnlHora.Caption = Format$(Now, "HH:MM")
End Sub

