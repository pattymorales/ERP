VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Begin VB.Form FBuscarCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Clientes"
   ClientHeight    =   4845
   ClientLeft      =   285
   ClientTop       =   1545
   ClientWidth     =   8910
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4845
   ScaleWidth      =   8910
   Begin MSGrid.Grid grdResultados 
      Height          =   3210
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   7920
      _Version        =   65536
      _ExtentX        =   13971
      _ExtentY        =   5662
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdSalir 
      Height          =   795
      Left            =   7995
      TabIndex        =   12
      Top             =   3990
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Salir"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FBUSCLI.frx":0000
   End
   Begin Threed.SSCommand cmdEscoger 
      Height          =   795
      Left            =   7995
      TabIndex        =   11
      Tag             =   "1190"
      Top             =   3180
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Escoger"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FBUSCLI.frx":08DA
   End
   Begin Threed.SSCommand cmdBuscar 
      Height          =   795
      Index           =   1
      Left            =   7995
      TabIndex        =   10
      Tag             =   "1182;1287"
      Top             =   2370
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Si&guiente"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Picture         =   "FBUSCLI.frx":11B4
   End
   Begin Threed.SSCommand cmdBuscar 
      Height          =   795
      Index           =   0
      Left            =   7995
      TabIndex        =   9
      Tag             =   "1182;1287"
      Top             =   1560
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Buscar"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FBUSCLI.frx":1A8E
   End
   Begin Threed.SSFrame fraCliente 
      Height          =   990
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   4185
      _Version        =   65536
      _ExtentX        =   7382
      _ExtentY        =   1746
      _StockProps     =   14
      Caption         =   "Tipo de Cliente"
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optCliente 
         Height          =   225
         Index           =   0
         Left            =   1335
         TabIndex        =   1
         Top             =   360
         Width           =   1080
         _Version        =   65536
         _ExtentX        =   1905
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "&Persona"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optCliente 
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1932
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "&Compañía"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame fraCriterio 
      Height          =   990
      Left            =   4320
      TabIndex        =   13
      Top             =   15
      Width           =   4590
      _Version        =   65536
      _ExtentX        =   8096
      _ExtentY        =   1746
      _StockProps     =   14
      Caption         =   "Criterios de Búsqueda"
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   1
      ShadowStyle     =   1
      Begin Threed.SSOption optCriterio 
         Height          =   225
         Index           =   3
         Left            =   195
         TabIndex        =   2
         Top             =   225
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "&Alfabético"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optCriterio 
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   3
         Top             =   465
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Céd&ula o RUC"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame FraValorBusq 
      Height          =   540
      Left            =   60
      TabIndex        =   16
      Top             =   1005
      Width           =   7920
      _Version        =   65536
      _ExtentX        =   13971
      _ExtentY        =   953
      _StockProps     =   14
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   1
      ShadowStyle     =   1
      Begin VB.TextBox txtValor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         TabIndex        =   4
         Top             =   165
         Width           =   5940
      End
      Begin Threed.SSPanel lblEtiqueta 
         Height          =   255
         Left            =   75
         TabIndex        =   21
         Top             =   165
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Valor de Búsqueda:"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         ShadowColor     =   1
         Alignment       =   1
      End
   End
   Begin Threed.SSFrame FraAlfabetico 
      Height          =   540
      Left            =   60
      TabIndex        =   17
      Top             =   1005
      Width           =   7920
      _Version        =   65536
      _ExtentX        =   13971
      _ExtentY        =   953
      _StockProps     =   14
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   1
      ShadowStyle     =   1
      Begin VB.TextBox txtNombres 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6390
         MaxLength       =   32
         TabIndex        =   7
         Top             =   165
         Width           =   1485
      End
      Begin Threed.SSPanel lblNombres 
         Height          =   255
         Left            =   5460
         TabIndex        =   20
         Top             =   165
         Width           =   900
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Nombres :"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         ShadowColor     =   1
         Alignment       =   1
      End
      Begin VB.TextBox txtSApellido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3525
         MaxLength       =   16
         TabIndex        =   6
         Top             =   165
         Width           =   1890
      End
      Begin Threed.SSPanel lblMaterno 
         Height          =   255
         Left            =   2700
         TabIndex        =   19
         Top             =   165
         Width           =   825
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Materno :"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         ShadowColor     =   1
         Alignment       =   1
      End
      Begin VB.TextBox txtPApellido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   915
         MaxLength       =   16
         TabIndex        =   5
         Top             =   165
         Width           =   1755
      End
      Begin Threed.SSPanel lblPaterno 
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   165
         Width           =   825
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Paterno :"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         ShadowColor     =   1
         Alignment       =   1
      End
   End
   Begin VB.Label il_titulo02 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   15
      TabIndex        =   15
      Top             =   3780
      Visible         =   0   'False
      Width           =   9420
   End
End
Attribute VB_Name = "FBuscarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables Locales
Dim VLCriterio%, VLCriterio_soc%, VLCliente%
Dim VLNumLin As Integer

Public Sub PLTomarDatosEnvio(parOperacion As String)
'Procedimiento para tomar datos que van a registrarse del ente
Dim cadena As String
ReDim vtCliente(6) As String
    VGEnteP.cedula = ""
    VGEnteP.Nombres = ""
    VGEnteP.apellidop = ""
    VGEnteP.apellidom = ""
    If optCliente(0).Value = True Then 'Si es persona
        grdResultados.Col = 4
        VGEnteP.Nombres = grdResultados.Text
        grdResultados.Col = 2
        VGEnteP.apellidop = grdResultados.Text
        grdResultados.Col = 3
        VGEnteP.apellidom = grdResultados.Text
        grdResultados.Col = 5
        VGEnteP.cedula = grdResultados.Text
        If VGEnteP.cedula <> "" Then
            VGEnteP.TipoId = "01"
            VGEnteP.Pais = 14
        Else
            VGEnteP.TipoId = "02" 'Pasaporte
            grdResultados.Col = 6
            VGEnteP.cedula = grdResultados.Text
        End If
    Else 'Si es Compania
         grdResultados.Col = 2
         VGEnteP.Nombres = grdResultados.Text
         VGEnteP.TipoId = "03"
         VGEnteP.Pais = 14
         grdResultados.Col = 3
         VGEnteP.cedula = grdResultados.Text
    End If
End Sub

Private Sub cmdBuscar_click(Index As Integer)
Dim VTCriterio$
Dim VTModo%
Dim VLEnte$
VTsp$ = "sp_se_ente"

'If optCliente(0).Value Then 'persona
'   PMPasoValores SqlConn&, "@i_subtipo", 0, SQLCHAR&, "P"
'   PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1182"
'Else
'   If optCliente(1).Value Then 'compañia
'   PMPasoValores SqlConn&, "@i_subtipo", 0, SQLCHAR&, "C"
'   PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1182"
'   Else 'IFI
'    If optCliente(2).Value Then 'ifi
'        PMPasoValores SqlConn&, "@i_subtipo", 0, SQLCHAR&, "I"
'        PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1287"
'        VTsp$ = "sp_se_ente2"
'    Else 'OFI
'        PMPasoValores SqlConn&, "@i_subtipo", 0, SQLCHAR&, "E"
'        PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1182"
'    End If
'   End If
'End If
'Determinar Criterio de Búsqueda -> parámetro @i_tipo
'If optCriterio(0).Value Then 'Codigo
'   PMPasoValores SqlConn&, "@i_tipo", 0, SQLINT1&, "1"
'   VTCriterio$ = "C"
'Else
'   If optCriterio(1).Value Then 'Cedula/Ruc
'      PMPasoValores SqlConn&, "@i_tipo", 0, SQLINT1&, "2"
'      VTCriterio$ = "R"
'   Else
'      If optCriterio(2).Value Then 'Pasaporte
'         PMPasoValores SqlConn&, "@i_tipo", 0, SQLINT1&, "3"
'         VTCriterio$ = "P"
'      Else                            'Alfabetico
'         If optCliente(0).Value Then  'Persona
'            PMPasoValores SqlConn&, "@i_tipo", 0, SQLINT1&, "4"
'            PMPasoValores SqlConn&, "@i_s_apellido_s", 0, SQLVARCHAR&, txtSApellido.Text
'            PMPasoValores SqlConn&, "@i_nombre_s", 0, SQLVARCHAR&, txtNombres.Text
'            PMPasoValores SqlConn&, "@i_p_apellido_s", 0, SQLVARCHAR&, txtPApellido.Text
'            VTCriterio$ = "A"
'         Else
'            PMPasoValores SqlConn&, "@i_tipo", 0, SQLINT1&, "4"
'            VTCriterio$ = "A"
'         End If
'      End If
'   End If
'End If
'Buscar los 20 primeros o los siguientes
Select Case Index
Case 0      '20 primeros
    VTModo% = False
    VLNumLin = 0
    'If optTipBusq(1).Value Then
    '    PMPasoValores SqlConn&, "@i_modo", 0, SQLINT1&, "2"     'Busqueda especifica
    'Else
    '    PMPasoValores SqlConn&, "@i_modo", 0, SQLINT1&, "0"     'Busqueda Parcial
    'End If
    'Determinar Valor de búsqueda
    If optCriterio(3).Value Then 'Alfabetico
       If optCliente(0).Value Then 'persona
               'nombrelargo$ = ""
               'nombrelargo$ = txtPapellido.Text + Chr(32) + txtSapellido.Text + Chr(32) + txtNombres.Text
               'If nombrelargo$ = "  " Then
               '   nombrelargo$ = " "
               'End If
               'nombrelargo$ = RTrim$(nombrelargo$)
               'If nombrelargo$ = "" Then
               '   nombrelargo$ = " "
               'End If
               'PMPasoValores SqlConn%, "@i_nombre", 0, SQLVARCHAR%, (nombrelargo$)

     Else  'compania
         If txtValor.Text <> "" Then
            'PMPasoValores SqlConn&, "@i_nombre_s", 0, SQLCHAR&, (txtValor.Text)
         Else
            'PMPasoValores SqlConn&, "@i_nombre_s", 0, SQLCHAR&, "[A-Z]"
         End If

       End If
    End If
    If optCriterio(1).Value Then 'Ruc
        If txtValor.Text <> "" Then
            'PMPasoValores SqlConn&, "@i_ced_ruc_s", 0, SQLCHAR&, (txtValor.Text)
        Else
            'PMPasoValores SqlConn&, "@i_ced_ruc_s", 0, SQLCHAR&, "[0-9]"
        End If
    End If
Case 1      'siguientes
    VTModo% = True
    'PMPasoValores SqlConn&, "@i_modo", 0, SQLINT1&, "1"
    grdResultados.Row = grdResultados.Rows - 1
    VTTope& = grdResultados.Row
    Select Case VTCriterio$
        Case "A"
            grdResultados.Col = 2
            'If optCliente(0).Value Then 'persona
            '   PMPasoValores SqlConn&, "@i_p_apellido", 0, SQLVARCHAR&, (grdResultados.Text)
            '   grdResultados.Col = 3
            '   PMPasoValores SqlConn&, "@i_s_apellido", 0, SQLVARCHAR&, (grdResultados.Text)
            '   grdResultados.Col = 4
            '   PMPasoValores SqlConn&, "@i_nombre", 0, SQLVARCHAR&, (grdResultados.Text)
            '   PMPasoValores SqlConn&, "@i_s_apellido_s", 0, SQLVARCHAR&, txtSApellido.Text
            '   PMPasoValores SqlConn&, "@i_nombre_s", 0, SQLVARCHAR&, txtNombres.Text
            '   PMPasoValores SqlConn&, "@i_p_apellido_s", 0, SQLVARCHAR&, txtPApellido.Text

            'Else
            '   PMPasoValores SqlConn&, "@i_nombre_s", 0, SQLCHAR&, (txtValor.Text)
            '   PMPasoValores SqlConn&, "@i_nombre", 0, SQLVARCHAR&, (grdResultados.Text)
            'End If
        Case "C"     'codigo
            grdResultados.Col = 1
            'PMPasoValores SqlConn&, "@i_ente_s", 0, SQLINT4&, txtValor.Text
            'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (grdResultados.Text)
        Case "R"     'cedula o ruc
            If optCliente(0).Value Then 'persona
            '   PMPasoValores SqlConn&, "@i_ced_ruc_s", 0, SQLCHAR&, (txtValor.Text)
            '   grdResultados.Col = 5
            Else  'compañia
            '   grdResultados.Col = 3
            End If
            'PMPasoValores SqlConn&, "@i_ced_ruc", 0, SQLCHAR&, (grdResultados.Text)
            'PMPasoValores SqlConn&, "@i_ced_ruc_s", 0, SQLCHAR&, (txtValor.Text)
        Case "P"
            'grdResultados.Col = 4
            'PMPasoValores SqlConn&, "@i_pasaporte_s", 0, SQLCHAR&, (txtValor.Text)
            'PMPasoValores SqlConn&, "@i_pasaporte", 0, SQLCHAR&, (grdResultados.Text)
        'Case "T"
        '    grdResultados.Col = 3
        '    PMPasoValores SqlConn%, "@i_tipo_soc_hecho", 0, SQLCHAR%, (grdResultados.Text)
        End Select
    End Select
    'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", VTsp$, True, "Búsqueda de Clientes") Then
    '  PMMapeaGrid SqlConn&, grdResultados, VTModo%
    '  PMChequea SqlConn&
    '   If Val(grdResultados.Tag) >= VGMaximoRows% - 1 Then
    '          If (grdResultados.Rows + VGMaximoRows%) > 16000 Then
    '            'Enviar un message box de aviso de limite
    '            MsgBox "Límite de Consulta excedido ", 64, "Busqueda Alfabética"
    '            cmdBuscar(1).Enabled = False
    '          Else
    '              cmdBuscar(1).Enabled = True
    '              If VTModo% = True Then
    '                If grdResultados.Rows > VGMaximoRows% Then
    '                    grdResultados.TopRow = VTTope&   'grdValores.Rows - MaximoRows% - 2
    '                End If
    '              End If
    '          End If
    '   Else
    '        If VTModo% = True Then
    '             If grdResultados.Tag > 0 Then
    '                grdResultados.TopRow = grdResultados.Rows - VGMaximoRows%
    '             End If
    '        End If
    '        cmdBuscar(1).Enabled = False
    '   End If
    'End If
End Sub

Private Sub cmdBuscar_GotFocus(Index As Integer)
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  Index       identifica al botón
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    Select Case Index
    Case 0
        FPrincipal!pnlHelpLine.Caption = "Buscar Clientes de acuerdo a especificaciones"
    Case 1
        FPrincipal!pnlHelpLine.Caption = "Buscar el siguiente grupo de Clientes"
    End Select
End Sub

Private Sub cmdEscoger_Click()
'*********************************************************
'Objetivo:  Esta rutina guarda el contenido de la línea
'           del grid escogida en el arreglo global VGBusqueda()
'           del cual, al salir de la forma, se pueden
'           extraer los datos de un cliente.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
Dim VTOpe As String
Dim VTRow%
    'Almacenar la línea escogida en el grid
    VTRow% = grdResultados.Row
    'Verificar que el grid contenga datos
    grdResultados.Col = 1
    grdResultados.Row = 1
    If grdResultados.Rows >= 2 And grdResultados.Text <> "" Then
        grdResultados.Row = VTRow%
        VTCols% = grdResultados.Cols - 1

        ' Chequeo realizado para salvaguardarse de la posibilidad de que el ente haya
        ' sido eliminado (desde esta sesion u otra sesion), mientras esta pantalla estuvo activa
    
        ReDim VTArreglo(5) As String
        grdResultados.Col = 1
        VLCodEnte$ = Trim$(grdResultados.Text)
        If Val(VLCodEnte$) <> 0 Then
            'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (VLCodEnte$)
            'codigo de transaccion
            'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1190"
            'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_qrente", True, " Cod. Cliente") Then
            '    vtr% = FMMapeaArreglo(SqlConn&, vtarreglo())
            '    PMChequea SqlConn&
            'Else
            '    Exit Sub
            'End If
        End If
        VGTipoCliente = VTArreglo(2)
        ReDim VGBusqueda(0 To VTCols%)
    'SGU se implementa validacion de clientes en listas reservadas
    If VGValida = True Then
        If optCliente(0).Value = True Or optCliente(1).Value = True Then
            VTOpe = "U"
            VGNumP = 2
            'Inicializar Estructuras de Validacion
            ReDim VGValidacionCt(4)
            ReDim VGValidacionRelCt(4)
            VGEnteCt.NumId = ""
            VGEnteCt.TipoId = ""
            VGEnteCt.Pais = ""
            VGEnteCt.apellidop = ""
            VGEnteCt.apellidom = ""
            VGEnteCt.nombre = ""
            VGEnteCt.CambioInfo = ""
            'Tomar datos de envio
            PLTomarDatosEnvio VTOpe
        End If
    End If
   'Fin SGU
        If optCliente(0).Value = True Then 'persona
            VGBusqueda(0) = "P"
            'PGA
            grdResultados.Col = 1
            VGEnte$ = grdResultados.Text
            grdResultados.Col = 2
            temporal$ = grdResultados.Text
            grdResultados.Col = 3
            temporal$ = temporal$ + " " + grdResultados.Text
            grdResultados.Col = 4
            temporal$ = temporal$ + " " + grdResultados.Text
            VGNombre$ = temporal$
            VGTipoDesc$ = "BENEFICIARIO FINAL"
        Else
            If optCliente(1).Value = True Then 'persona
                VGBusqueda(0) = "C" 'compañía
                'PGA
                grdResultados.Col = 1
                VGEnte$ = grdResultados.Text
                grdResultados.Col = 2
                VGNombre$ = grdResultados.Text
                VGTipoDesc$ = "BENEFICIARIO FINAL"
            Else
                VGBusqueda(0) = "I" 'IFI
                grdResultados.Col = 1
                VGEnte$ = grdResultados.Text
                grdResultados.Col = 2
                VGNombre$ = grdResultados.Text
                'PGA
                VGTipoDesc$ = "IFI"
            End If
        End If
        
        For i% = 1 To grdResultados.Cols - 1
            grdResultados.Col = i%
            VGBusqueda(i%) = grdResultados.Text
        Next i%
        FPrincipal.pnlHelpLine.Caption = ""
        FPrincipal.pnlTransaccionLine.Caption = ""
        If FBuscarCliente.Tag <> "C" Then
           'FBuscarCliente.Hide
            Unload FBuscarCliente 'DMO no utilizo el salir por que se borran las variables globales
        End If
    End If
End Sub

Private Sub cmdEscoger_GotFocus()
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlHelpLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Escoger un Cliente"
End Sub

Private Sub cmdSalir_Click()
'*********************************************************
'Objetivo:  Inicializa el arreglo global VGBusqueda y ocul-
'           ta la forma. (No la descarga)
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    ReDim VGBusqueda(10)
    Unload FBuscarCliente 'DMO descargo la forma
    'FBuscarCliente.Hide
End Sub

Private Sub cmdSalir_GotFocus()
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlHelpLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Salir de la Forma de Búsqueda"
End Sub

Private Sub Form_Load()
    'optcriterio(2).Enabled = False
   If optCliente(0).Value = True Then
    FraValorBusq.Visible = False
    FraAlfabetico.Visible = True
    FraValorBusq.Enabled = False
    FraAlfabetico.Enabled = True
   Else
   'If optcliente(1).Value = True And optcriterio(3).Value = True Then
    FraValorBusq.Visible = True
    FraAlfabetico.Visible = False
    FraValorBusq.Enabled = True
    FraAlfabetico.Enabled = False
   'End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*********************************************************
'Objetivo:  Limpia las líneas de Mensaje y Transaccion
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal.pnlHelpLine.Caption = ""
    FPrincipal.pnlTransaccionLine.Caption = ""
    VLCriterio% = 0      'criterio de busq. para persona/sociedad
    VLCriterio_soc% = 0  'criterio de busq. para sociedad de hecho
    VLNumLin = 0
    VGValida = False
End Sub

Private Sub grdResultados_Click()
'*********************************************************
'Objetivo:  Selecciona una línea del grid
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    PMLineaG grdResultados
End Sub

Private Sub grdResultados_DblClick()
'*********************************************************
'Objetivo:  Llama a la rutina del botón escoger
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    cmdEscoger_Click
End Sub

Private Sub grdResultados_GotFocus()
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    If FBuscarCliente.Tag = "C" Then 'entro por el menu de consulta y no de busqueda
       FPrincipal!pnlHelpLine.Caption = ""
    Else
       FPrincipal!pnlHelpLine.Caption = "Doble-click para Escoger un Cliente"
    End If
End Sub

Private Sub optCliente_Click(Index As Integer, Value As Integer)
'*********************************************************
'Objetivo:  Sirve para seleccionar el tipo de cliente a
'           buscar: Personas o Compañías. Según el tipo
'           de cliente habilita los criterios apropiados.
'           Cuando el tipo de cliente cambia, limpia el
'           grid de resultados
'Input   :  Index       identifica al option buttom
'                       0 Persona / 1 Compañía
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    Select Case Index
    Case 0
        'Persona
        If Value = True Then
            If Value = True Then
            'If VLCliente% = 1 Then 'estaba seleccionada compañía
                VLCliente% = 0
                PMLimpiaG grdResultados
                cmdBuscar(1).Enabled = False
            End If
            optCriterio(1).Caption = "Céd&ula o RUC"
            If optCliente(0).Value = True And optCriterio(3).Value = True Then
               FraValorBusq.Visible = False
               FraAlfabetico.Visible = True
               FraValorBusq.Enabled = False
               FraAlfabetico.Enabled = True
            End If
        End If
    Case 1
        'Compañía
        If Value = True Then
            If Value = True Then
            'If VLCliente% = 0 Then 'estaba seleccionada persona
                VLCliente% = 1
                PMLimpiaG grdResultados
                cmdBuscar(1).Enabled = False
            End If
            optCriterio(1).Caption = "Céd&ula o RUC"
            FraValorBusq.Visible = True
            FraAlfabetico.Visible = False
            FraValorBusq.Enabled = True
            FraAlfabetico.Enabled = False
            optCriterio(3).Value = True
        End If
    Case 2
       'Institución Financiera
       optCriterio(2).Enabled = True
       optCriterio(2).Caption = "&Codigo Ifi"
       optCriterio(1).Caption = "&Ruc"
       If Value = True Then
           VLCliente% = 2
           PMLimpiaG grdResultados
           cmdBuscar(1).Enabled = False
       End If
       FraValorBusq.Visible = True
       FraAlfabetico.Visible = False
       FraValorBusq.Enabled = True
       FraAlfabetico.Enabled = False

    Case 3
       'OFI
       optCriterio(2).Enabled = True
       optCriterio(2).Caption = "&Codigo Ofi"
       optCriterio(1).Caption = "&Ruc"
        If Value = True Then
           VLCliente% = 3
           PMLimpiaG grdResultados
           cmdBuscar(1).Enabled = False
        End If
    FraValorBusq.Visible = True
    FraAlfabetico.Visible = False
    FraValorBusq.Enabled = True
    FraAlfabetico.Enabled = False

    End Select
End Sub

Private Sub optCliente_GotFocus(Index As Integer)
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  Index       identifica al option buttom
'                       0 Persona / 1 Compañía
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Seleccionar el Tipo de Cliente que se desea Buscar"
End Sub

Private Sub optCriterio_Click(Index As Integer, Value As Integer)
'*********************************************************
'Objetivo:  Nos permite escoger un criterio de búsqueda
'           de clientes. Cuando el criterio cambia, ini-
'           cializa el campo con el valor inicial a "%"
'Input   :  Index       identifica al option buttom.
'                       0 Código | 1 Cedula o RUC |
'                       2 Pasaporte | 3 Alfabético
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    txtNombres.Text = ""
    txtPApellido = ""
    txtSApellido = ""
    txtValor = ""
    Select Case Index
    Case 0
        'Código
        If Value = True Then
            If optCriterio(0).Enabled = True Then
            ' If VLCriterio% <> 0 Then 'estaba seleccionado otro criterio
                VLCriterio% = 0
                'txtvalor.Text = "%"
                FraValorBusq.Visible = True
                FraAlfabetico.Visible = False
                FraValorBusq.Enabled = True
                FraAlfabetico.Enabled = False
                cmdBuscar(1).Enabled = False
            End If
        End If
    Case 1
        'Cédula/Ruc
        If Value = True Then
            If VLCriterio% <> 1 Then 'estaba seleccionado otro criterio
                VLCriterio% = 1
                'txtvalor.Text = "%"
                FraValorBusq.Visible = True
                FraAlfabetico.Visible = False
               FraValorBusq.Enabled = True
                FraAlfabetico.Enabled = False
                cmdBuscar(1).Enabled = False
            End If
        End If
    Case 2
        'Pasaporte
        If Value = True Then
            If VLCriterio% <> 2 Then 'estaba seleccionado otro criterio
                VLCriterio% = 2
                'txtvalor.Text = "%"
                FraValorBusq.Visible = True
                FraAlfabetico.Visible = False
                FraValorBusq.Enabled = True
                FraAlfabetico.Enabled = False
                cmdBuscar(1).Enabled = False
            End If
        End If
    Case 3
        'Alfabético
        If Value = True Then
            If VLCriterio% <> 3 Then 'estaba seleccionado otro criterio
                VLCriterio% = 3
                 If optCliente(1).Value = True Then
                    FraValorBusq.Visible = True
                    FraAlfabetico.Visible = False
                    FraValorBusq.Enabled = True
                    FraAlfabetico.Enabled = False
                 Else
                    FraValorBusq.Visible = False
                    FraAlfabetico.Visible = True
                    FraValorBusq.Enabled = False
                    FraAlfabetico.Enabled = True
                 End If
                    cmdBuscar(1).Enabled = False
              End If
        End If
    End Select
End Sub

Private Sub optCriterio_GotFocus(Index As Integer)
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  Index       identifica al option buttom
'                       0 Codigo | 1 Cedula 0 RUC
'                       2 Pasaporte | 3 Alfabético
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Seleccionar el Criterio de Búsqueda"
End Sub

Private Sub optcriterio_soc_Click(Index As Integer, Value As Integer)
    Select Case Index
    Case 0
        'Código
        If Value = True Then
            If VLCriterio_soc% <> 0 Then 'estaba seleccionado otro criterio
               VLCriterio_soc% = 0
               'txtvalor.Text = "%"
               cmdBuscar(1).Enabled = False
            End If
        End If
    Case 1
        'Cédula/Ruc
        If Value = True Then
            If VLCriterio_soc% <> 1 Then 'estaba seleccionado otro criterio
                VLCriterio_soc% = 1
                'txtvalor.Text = "%"
                cmdBuscar(1).Enabled = False
            End If
        End If
    Case 2
        'codigo de sociedad
        If Value = True Then
            If VLCriterio_soc% <> 2 Then 'estaba seleccionado otro criterio
                VLCriterio_soc% = 2
                'txtvalor.Text = "%"
                cmdBuscar(1).Enabled = False
            End If
        End If
    End Select

End Sub

Private Sub optTipBusq_Click(Index As Integer, Value As Integer)
    Select Case Index
    Case 0
    Case 1
         cmdBuscar(1).Enabled = False
    End Select
End Sub

Private Sub optTipBusq_GotFocus(Index As Integer)
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Seleccionar el Tipo de Busqueda"
End Sub

Private Sub txtNombres_Change()
    'If (txtSapellido.Text = "" Or txtPapellido.Text = "") And txtNombres.Text <> "" Then
    '   txtPapellido.SetFocus
    '   If txtPApellido.Text = "" Then
    '        MsgBox "Ingrese Apellido Paterno ", 48, "Busqueda Alfabética"
    '   Else
    '       If txtSApellido.Text = "" Then
    '            MsgBox "Ingrese Apellido Materno ", 48, "Busqueda Alfabética"
    '       End If
    '   End If
    'End If
   
End Sub

Private Sub txtNombres_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Ingreso de Nombres"
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    If optCliente(0).Value And optCriterio(3).Value Then
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% <> 39) And (KeyAscii% <> 241) And (KeyAscii% <> 209) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If
End Sub

Private Sub txtNombres_LostFocus()
txtNombres.Text = Trim$(UCase$((txtNombres.Text)))
End Sub

Private Sub txtPApellido_Change()
    txtSApellido.Text = ""
    txtNombres.Text = ""
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Ingreso de Apellido Paterno"
End Sub

Private Sub txtPApellido_KeyPress(KeyAscii As Integer)
    If optCliente(0).Value And optCriterio(3).Value Then
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% <> 39) And (KeyAscii% <> 241) And (KeyAscii% <> 209) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If
End Sub

Private Sub txtPApellido_LostFocus()
txtPApellido.Text = Trim$(UCase$((txtPApellido.Text)))
End Sub

Private Sub txtSApellido_Change()
    'If txtPApellido.Text = "" And txtSApellido.Text <> "" Then
    '   txtPApellido.SetFocus
    '   txtSApellido.Text = ""
    '   MsgBox "Ingrese Apellido Paterno ", 48, "Busqueda Alfabética"
    'Else
        'txtSapellido.Text = ""
        'txtNombres.Text = ""
        'FPrincipal!pnlTransaccionLine.Caption = ""
        FPrincipal!pnlHelpLine.Caption = "Ingreso de Apellido Materno"
    'End If
End Sub

Private Sub txtSApellido_KeyPress(KeyAscii As Integer)
    If optCliente(0).Value And optCriterio(3).Value Then
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% <> 39) And (KeyAscii% <> 241) And (KeyAscii% <> 209) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If

End Sub

Private Sub txtSApellido_LostFocus()
    'If txtSapellido = "" And txtPapellido = "" Then
    '   txtPapellido.SetFocus
    '   MsgBox "Ingrese Apellido Paterno ", 48, "Busqueda Alfabética"
    'Else
       txtSApellido.Text = Trim$(UCase$((txtSApellido.Text)))
    'End If
End Sub

Private Sub txtValor_GotFocus()
'*********************************************************
'Objetivo:  Despliega el Mensaje de Ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = "Dar el valor inicial de Búsqueda"
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea alfanumerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'22/Ene/94  M.Davila            Emisión Inicial
'********************************************************
If optCliente(0).Value And optCriterio(3).Value Then
    If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% <> 39) And (KeyAscii% <> 241) And (KeyAscii% <> 209) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
        KeyAscii% = 0
    Else
        KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        'keyAscii% = Asc(Chr$(keyAscii%))
    End If
Else
    If optCriterio(1).Value Then
        If (KeyAscii% <> 8) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    Else
      If optCriterio(2).Value Then
              If (KeyAscii% <> 8) And (KeyAscii% <> 45) And (KeyAscii% <> 241) And (KeyAscii% <> 209) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
                KeyAscii% = 0
              Else
                KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
              End If
        Else
            If optCriterio(0).Value Then
                If (KeyAscii% <> 8) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
                    KeyAscii% = 0
                Else
                    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
                End If
            Else
                KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
            End If
        End If
    End If
End If
End Sub

Private Sub txtValor_LostFocus()
'*********************************************************
'Objetivo:  Elimina blancos a izquierda y derecha
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    txtValor.Text = Trim$(UCase$((txtValor.Text)))
End Sub


