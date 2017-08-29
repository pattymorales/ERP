VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FApGeneral 
   Appearance      =   0  'Flat
   Caption         =   "Datos de Referencia"
   ClientHeight    =   6495
   ClientLeft      =   1200
   ClientTop       =   450
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6495
   ScaleWidth      =   9330
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   7410
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   19
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
         Picture         =   "FAPGNRAL.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   20
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
         Picture         =   "FAPGNRAL.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   21
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
         Picture         =   "FAPGNRAL.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   22
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
         Picture         =   "FAPGNRAL.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
         TabIndex        =   23
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
         Picture         =   "FAPGNRAL.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin Threed.SSFrame fraDireccion 
      Height          =   1545
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   8355
      _Version        =   65536
      _ExtentX        =   14737
      _ExtentY        =   2725
      _StockProps     =   14
      Caption         =   "Dirección"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin MSMask.MaskEdBox txtTelefono 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   1080
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDescripcionDireccion 
         Height          =   585
         Left            =   1800
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   480
         Width           =   3750
      End
      Begin VB.Label lblNumDireccion 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1800
         TabIndex        =   17
         Top             =   195
         Width           =   735
      End
      Begin VB.Label lblEtiqueta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "No. de Dirección:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Width           =   1515
      End
      Begin VB.Label lblEtiqueta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   5
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "No.Telefónico"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label lblEtiqueta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Lista de Direcciones"
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
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   9300
      Begin MSGrid.Grid grdLista 
         Height          =   915
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   1614
         _StockProps     =   77
         ForeColor       =   0
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
      Begin Threed.SSCommand cmdEscoger 
         Height          =   405
         Left            =   8350
         TabIndex        =   15
         Top             =   360
         Width           =   885
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Esco&ger"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdEliminar 
         Height          =   405
         Left            =   8350
         TabIndex        =   16
         Top             =   800
         Width           =   885
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Eliminar"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos Generales del Cliente"
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
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9300
      Begin Threed.SSPanel pnlCedRuc 
         Height          =   285
         Left            =   2295
         TabIndex        =   8
         Top             =   525
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   503
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin Threed.SSPanel pnlNombre 
         Height          =   600
         Left            =   4080
         TabIndex        =   9
         Top             =   195
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   1058
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel pnlEnte 
         Height          =   300
         Left            =   2295
         TabIndex        =   10
         Top             =   195
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   529
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin VB.Label lblDatos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "No. de Identificación:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   250
         Width           =   2115
      End
      Begin VB.Label lblDatos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "CI./RUC/PASAPORTE:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   530
         Width           =   1995
      End
   End
End
Attribute VB_Name = "FApGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'   Archivo:        FAPGNRAL.FRM
'   Producto:       Master Information Subsystem
'   Diseñado por:   Myriam Dávila
'   Fecha de Documentación: 14/Ene/94
'*********************************************************
'                       IMPORTANTE
' Esta Aplicación es parte de los paquetes bancarios pro-
' piedad de MACOSA, representantes exclusivos para el Ecua-
' dor de la NCR CORPORATION.  Su uso no autorizado queda
' expresamente prohibido así como cualquier alteración o
' agregado hecho por alguno de sus usuarios sin el debido
' consentimiento por escrito de la Presidencia Ejecutiva
' de MACOSA o su representante
'*********************************************************
' Forma: FApGeneral
' Descripcion:  Esta Forma sirve para los atributos multi-
'               valor de clientes (Direcciones, Casillas
'               Postales, Empleos, Propiedades, Ref. Eco-
'               nómicas, Ref. Personales)
' VARIABLES GLOBALES:
'           ServerName$         nombre del servidor
'           VGOficina$          codigo de la oficina de conexión
' MODULOS:
'           MHELPG.BAS = Rutinas para usar la ventana de ayuda
' LLAMADA:
'           para llamar a esta Forma, se requiere poner en la
'           propiedad .Tag la operacion a realizar "I" (insertar)
'           o "U" (actualizar)
'*********************************************************

' Constantes
Const F5% = 116         'código de la tecla F5
Const IDYES = 6         'botón YES
Const IDNO = 7          'botón NO

' Variables Locales
Dim VLPaso%             'chequea cambios en un campo
Dim VLOperacion$        'operacion a realizarse en la bdd
Dim VLPrimeraVez%       'sirve para inicializar forma
Dim VLOficina$          'variable local con la oficina
Dim VLpais$             'default de pais para la casilla postal
Dim VLPasoM%            'chequea que cambio el campo sueldo o valor de la propiedad
Dim VLanterior$         'chequea que cambio el campo provincia para la creacion de postales
Dim VLTranOk$
Dim VLTranI$
Dim VLTranU$



Private Sub cmdBuscar_click()
'*********************************************************
'Objetivo:  Busca los registros existentes para un cliente
'           en la tabla apropiada, de acuerdo a la función
'           que esté realizando la forma actualmente (pro-
'           piedad .Tag de la forma)
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
Dim VTsp$
     VTsp$ = "sp_direccion_cons"
     'codigo de transaccion
     'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1227"
     'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "S"
     'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
     'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", VTsp$, True, "Buscar Registros Existentes") Then
        If VTsp$ = "sp_ref_eco2" Then
             'PMMapeaGrid SqlConn&, grdLista, False
             'PMMapeaGrid SqlConn&, grdLista, True
            ' PMMapeaGrid SqlConn&, grdLista, True
             'PMMapeaGrid SqlConn&, grdLista, True
        Else
            ' PMMapeaGrid SqlConn&, grdLista, False
        End If
        ' PMChequea SqlConn&
    'End If
    If VLTranOk$ = "S" Then
       FPrincipal!pnlTransaccionLine.Caption = "Tran. Ok."
    End If

End Sub

Private Sub cmdBuscar_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de Ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlHelpLine.Caption = " Buscar Registros Existentes"
End Sub

Private Sub cmdcompania_Click()
   FApRedCia.Show 1
End Sub

Private Sub cmdcompania_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Creación de Compañia "
End Sub

Private Sub cmdEliminar_Click()
'*********************************************************
'Objetivo:  Elimina un registro para un ente dado de la
'           tabla apropiada.  El registro a borrar se
'           toma de la línea seleccionada en el grid
'           con la lista de Registros existentes
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
Dim VTRow%, VTresp%, VTMsg$
Dim VTCod$
    VTRow% = grdLista.Row
    grdLista.Row = 1
    grdLista.Col = 1
    If grdLista.Text <> "" And grdLista.Rows >= 2 Then
        VTMsg$ = "Desea borrar el Registro?" + Chr(13)
        grdLista.Row = VTRow%
        grdLista.Col = 1
        VTMsg$ = VTMsg$ + " " + grdLista.Text + " "
        grdLista.Col = 2
        VTMsg$ = VTMsg$ + grdLista.Text
        VTresp% = MsgBox(VTMsg$, 36, FApGeneral.Tag)
        If VTresp% = IDYES Then
        VTbdd$ = "cobis"
        VTsp$ = "sp_direccion_dml"
        grdLista.Col = 1
        'PMPasoValores SqlConn&, "@i_direccion", 0, SQLINT1&, (grdLista.Text)
        VTCod$ = grdLista.Text
        'codigo de transaccion
        'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1226"
        'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "D"
        'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
        'If FMTransmitirRPC(SqlConn&, ServerName$, VTbdd$, VTsp$, True, "Eliminar Registro") Then
            'PMChequea SqlConn&
            If lblNumDireccion.Caption = VTCod$ Then
               jmbLimpiar_Click
            End If
        
        End If
        'cmdBuscar_click
    End If
End Sub

Private Sub cmdEliminar_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlHelpLine.Caption = " Eliminar un Registro Señalado"
End Sub

Private Sub cmdEscoger_Click()
 cmdModificar_Click
End Sub




Private Sub cmdModificar_Click()
'*********************************************************
'Objetivo:  Escoge el registro señalado en la lista de
'           registros existentes para modificar, entonces
'           pasa el contenido de los campos del registro
'           a los campos de input del frame inferior co-
'           rrespondiente.  La operacion se pone en "U"
'           (Modificación)
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
ReDim VTArreglo(40) As String
Dim VTRow%
    jmbLimpiar_Click
    VTRow% = grdLista.Row
    On Error Resume Next
    grdLista.Row = 1
    grdLista.Col = 1
    If grdLista.Text <> "" And grdLista.Rows >= 2 Then
            'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "Q"
            'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
            grdLista.Row = VTRow%
            grdLista.Col = 1
            'PMPasoValores SqlConn&, "@i_direccion", 0, SQLINT1&, (grdLista.Text)
            'codigo de transaccion
            'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1228"
            'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_direccion_cons", True, "Dirección") Then
                 'VTR% = FMMapeaArreglo(SqlConn&, VTArreglo())
                 'PMMapeaListaH SqlConn&, lstTelefonos, False
                 'PMChequea SqlConn&
                 lblNumDireccion.Caption = VTArreglo(1)
                 txtDescripcionDireccion.Text = Trim$(VTArreglo(2))
                 txtProvincia_dir.Text = VTArreglo(4)    'PROVINCIA
                 lblProvincia_dir.Caption = VTArreglo(5) 'DESCRIPCION
                 txtParroquia.Text = VTArreglo(7)        'Parroquia
                 lblParroquia.Caption = VTArreglo(8)     'DESCRIPCION
                 txtTDireccion.Text = VTArreglo(9)       'TIPO DIRECCION
                 lblTDireccion.Caption = VTArreglo(10)    'DESCRIPCION
                 txtSector.Text = VTArreglo(11)           'SECTOR
                 lblSector.Caption = VTArreglo(12)       'DESCRIPCION
                 txtOficina.Text = VTArreglo(15)         'OFICINA
                 lblOficina.Caption = VTArreglo(16)      'DESCRIPCIOn
                 lblNumTelefonos.Caption = Str(lstTelefonos.ListCount)
                txtTelefono.Text = ""
                txtTTelefono.Text = ""
                lblTTelefono.Caption = ""
                'Inicio CS Req #506-2009 Direcciones de Correspondencia
                Select Case VTArreglo(23)
                            Case "S"
                                chKDirCorre.Value = 1
                            Case Else
                                chKDirCorre.Value = 0
                        End Select
                'Fin CS Req #506-2009 Direcciones de Correspondencia
                fraDireccion.Enabled = True
                fraDireccion.Caption = "Dirección [Actualización]"
            End If
        
        VLOperacion$ = "U"
        cmdTransmitir.Enabled = True
    'End If
End Sub

Private Sub cmdModificar_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Modificar un Registro"
End Sub








Private Sub cmdUpTelefono_Click()
'*********************************************************
'Objetivo:  Modifica un teléfono de una dirección. Si la
'           operación es "I"(inserción) entonces la modifica-
'           ción es local, si la operacion es "U" (modifi-
'           cación) primero se modifica la bdd y luego
'           en forma local
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
  If txtTelefono.Text <> "" And txtTTelefono.Text <> "" And txtTelefono.Tag <> "" Then
    If lstTelefonos.ListCount > 0 And lstTelefonos.ListIndex >= 0 Then
        i% = Len(txtTelefono.Text)
        If i% < 12 Then
           For j% = 1 To 12 - i%
               txtTelefono.Text = txtTelefono.Text + " "
           Next j%
        End If

        Select Case VLOperacion$
        Case "I"
            'En el caso que no existan todos los datos de direccion actulizo tan solo el
            'telefono antes de ingresar los nuevos valores.

            lstTelefonos.ListIndex = txtTelefono.Tag
            lstTelefonos.RemoveItem txtTelefono.Tag
            lstTelefonos.AddItem txtTelefono.Text + Chr(9) + txtTTelefono.Text + Chr(9) + lblTTelefono.Caption, txtTelefono.Tag
            txtTelefono.Text = ""
            txtTTelefono.Text = ""
            lblTTelefono.Caption = ""
            txtTelefono.Tag = ""
        
        Case "U"
            'En el caso que existan todos los datos de direccion actualizo

             'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "U"
             'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
             'PMPasoValores SqlConn&, "@i_direccion", 0, SQLINT1&, (lblNumDireccion.Caption)
             'PMPasoValores SqlConn&, "@i_valor", 0, SQLVARCHAR&, (txtTelefono.Text)
             'PMPasoValores SqlConn&, "@i_tipo_telefono", 0, SQLCHAR&, (txtTTelefono.Text)
             'PMPasoValores SqlConn&, "@i_secuencial", 0, SQLINT1&, Str(Val(txtTelefono.Tag) + 1)

             'codigo de transaccion
             'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "112"
             'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_telefono", True, "Añadir Teléfono") Then
                'PMChequea SqlConn&
                lstTelefonos.ListIndex = txtTelefono.Tag
                lstTelefonos.RemoveItem txtTelefono.Tag
                lstTelefonos.AddItem txtTelefono.Text + Chr(9) + txtTTelefono.Text + Chr(9) + lblTTelefono.Caption, txtTelefono.Tag
                txtTelefono.Text = ""
                txtTTelefono.Text = ""
                lblTTelefono.Caption = ""
                txtTelefono.Tag = ""
            'End If
        End Select
    End If
  End If
End Sub

Private Sub cmdUpTelefono_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Modificar un número Telefónico"
End Sub
Private Sub chkEstatus_GotFocus(Index As Integer)
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Seleccione el Estatus"

End Sub
Private Sub chkTipoOperacion_GotFocus(Index As Integer)
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Seleccione un Tipo de Operación"
End Sub

Private Sub Form_Activate()
    If VLPrimeraVez% Then
        ' Con este Case se Maneja el tipo de atributo multivalor
        ' que se va a afectar
        FApGeneral.Caption = FApGeneral.Tag
        fraDetalle.Caption = "Lista de " + FApGeneral.Tag
        FApGeneral.HelpContextID = 152
        fraDireccion.Visible = True
        fraDireccion.Caption = "Dirección [Creación]"
        'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "H"
        'PMPasoValores SqlConn&, "@i_tipo_h", 0, SQLCHAR&, "V"
        'PMPasoValores SqlConn&, "@i_ciudad", 0, SQLINT4&, (txtCiudad(0).Text)
        'PMPasoValores SqlConn&, "@i_parroquia", 0, SQLINT4&, (txtParroquia.Text)
        'PMPasoValores SqlConn&, "@i_provincia", 0, SQLINT4&, (txtProvincia_dir.Text)
        'codigo de transaccion
        'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1559"
        ReDim VTArreglo(4) As String
        'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_parroquia", True, "Ayuda Sector") Then
        'VTR% = FMMapeaArreglo(SqlConn&, VTArreglo())
        'PMChequea SqlConn&
        'End If
        txtTelefono.MaxLength = 12
        cmdBuscar_click
        VLPrimeraVez% = False
        VLPaso% = True
    End If
End Sub

Private Sub Form_Load()
'*********************************************************
'Objetivo:  Setea las variables apropiadas para el inicio
'           de la operación de la forma.
'           VLPrimeraVez: es para el evento Activate
'           VLOperacion: es "I" (inserción)
'           VLOficina$: toma de la variable global corres-
'                       pondiente.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    VLPrimeraVez% = True
    VLTranOk$ = "N"
    VLOperacion$ = "I"
    VLOficina$ = VGOficina$
    VLFormato$ = "mm/dd/yyyy"
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
End Sub

Private Sub grdLista_Click()
'*********************************************************
'Objetivo:  Llama al procedimiento que resalta la línea
'           del grid sobre la cual se da click con el
'           mouse.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    PMLineaG grdLista
End Sub

Private Sub grdLista_DblClick()
'*********************************************************
'Objetivo:  Con doble click se indica que se quiere modi-
'           ficar el registro correspondiente. Se llama
'           a la rutina del botón Modificar.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    cmdModificar_Click
End Sub

Private Sub grdLista_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = fraDetalle.Caption + " Doble Click para Escoger"
End Sub

Private Sub lstTelefonos_DblClick()
'*********************************************************
'Objetivo:  Un doble click en un teléfono de la lista in-
'           dica que se lo quiere modificar. Se pasa el
'           contenido del item de la lista a los campos
'           superiores correspondientes.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
Dim pos1%, pos2%
    If lstTelefonos.ListCount > 0 Then
        txtTelefono.Tag = lstTelefonos.ListIndex
        If txtTelefono.Tag >= 0 Then
            pos1% = InStr(1, lstTelefonos.Text, Chr(9))
            pos2% = InStr(pos1% + 1, lstTelefonos.Text, Chr(9))
            txtTelefono.Text = Mid$(lstTelefonos.Text, 1, pos1% - 1)
            txtTTelefono.Text = Mid$(lstTelefonos.Text, pos1% + 1, pos2% - pos1% - 1)
            lblTTelefono.Caption = Mid$(lstTelefonos.Text, pos2% + 1, Len(lstTelefonos.Text))
            txtTelefono.SetFocus
        End If
    End If
End Sub

Private Sub lstTelefonos_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Doble Click para seleccionar un teléfono"
End Sub

Private Sub lstTelefonos_KeyPress(KeyAscii As Integer)
If KeyAscii% = 13 Then
    lstTelefonos_DblClick
End If
End Sub

Private Sub MaskedEdit1_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado e ini-
'           cializa la variable que chequea los cambios
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'28/Abr/94  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Fecha de Vencimiento Operación  [mm/dd/yyyy]"
    txtFechaVencimiento.SelStart = 0
    txtFechaVencimiento.SelLength = Len(txtFechaVencimiento.Text)

End Sub

Private Sub MaskedEdit1_LostFocus()
'*********************************************************
'Objetivo:  Verifica la validez de la fecha de acuerdo
'           al formato actual
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'28/Abr/94  M.Davila            Emisión Inicial
'*********************************************************
Dim VT%, vtd&
    On Error Resume Next
    If txtFechaVencimiento.ClipText <> "" Then
        VT% = FMVerFormato((txtFechaVencimiento.FormattedText), "mm/dd/yyyy")
        If Not VT% Then
            MsgBox "Formato de Fecha Inválido", 48, "Control Ingreso de Datos"
            txtFechaVencimiento.SetFocus
            Exit Sub
        Else
            vtd& = FMDateDiff("d", (txtFechaInicio.FormattedText), (txtFechaVencimiento.FormattedText), "mm/dd/yyyy")
            If vtd& <= 0 Then
                MsgBox "Fecha de Vencimiento debe ser MAYOR que Fecha de Inicio", 48, "Control Ingreso de Datos"
                txtFechaVencimiento.SetFocus
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub optTipoCasilla_GotFocus(Index As Integer)
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Seleccione un Tipo de Casilla"
End Sub


Private Sub Text4_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado e ini-
'           cializa la variable que chequea los cambios
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Observaciones a la Referencia"
    txtObservacionF.SelStart = 0
    txtObservacionF.SelLength = Len(txtObservacionF.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea alfanumerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))


End Sub

Private Sub Text4_LostFocus()
'*********************************************************
'Objetivo:  elimina blancos a izquierda y derecha
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    txtObservacionF.Text = Trim$(txtObservacionF.Text)

End Sub

Private Sub Text5_Change()
'*********************************************************
'Objetivo:  Determina que el campo calificacion ha cambiado.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    VLPaso% = False

End Sub

Private Sub Text5_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado e ini-
'           cializa la variable que chequea los cambios
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Calificación de la Referencia [F5] Ayuda"
    txtCalificacionF.SelStart = 0
    txtCalificacionF.SelLength = Len(txtCalificacionF.Text)

End Sub

Private Sub jmbIngresar_Click()
    Dim VTDirCorres$ 'CS Req #506-2009 Para escoger Dirección de Correspondencia
   'chequeo de campos
    If txtDescripcionDireccion.Text = "" Then
        MsgBox "Campo Dirección Mandatorio", 48, "Control Ingreso de Datos"
        txtDescripcionDireccion.SetFocus
        Exit Sub
    End If
    If txtTDireccion.Text = "" Then
        txtTDireccion.Text = "P"
    End If
        Select Case VLOperacion$
        Case "I"
            fraDireccion.Enabled = False
            cmdTransmitir.Enabled = False
            'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "I"
            'codigo de transaccion
            'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "109"
        Case "U"
             'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "U"
             'PMPasoValores SqlConn&, "@i_direccion", 0, SQLINT1&, (lblNumDireccion.Caption)
            'codigo de transaccion
             'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "110"
        End Select
         'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
         'PMPasoValores SqlConn&, "@i_descripcion", 0, SQLVARCHAR&, Trim$(txtDescripcionDireccion.Text)
         'PMPasoValores SqlConn&, "@i_tipo", 0, SQLVARCHAR&, (txtTDireccion.Text)
         'PMPasoValores SqlConn&, "@i_sector", 0, SQLVARCHAR&, (txtSector.Text)
         'PMPasoValores SqlConn&, "@i_provincia", 0, SQLINT4&, (txtProvincia_dir.Text)
         'PMPasoValores SqlConn&, "@i_verificado", 0, SQLCHAR&, "N"
         'PMPasoValores SqlConn&, "@i_parroquia", 0, SQLINT4&, (txtParroquia.Text)
         'PMPasoValores SqlConn&, "@i_ciudad", 0, SQLINT4&, (txtCiudad(0).Text)
         'Inicio CS Req #506-2009 Para escoger Direccion de Correspondencia
         If chKDirCorre.Value = 1 Then
            VTDirCorres$ = "S"
         Else
            VTDirCorres$ = "N"
         End If
         'PMPasoValores SqlConn&, "@i_dir_corre", 0, SQLCHAR&, VTDirCorres$
         'Fin CS Req #506-2009 Para escoger Direccion de Correspondencia
         'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_direccion_dml", True, "Dirección") Then
            If VLOperacion$ = "I" Then
                 'PMMapeaObjeto SqlConn&, lblNumDireccion
                 'PMChequea SqlConn&
                
                'Insertar Todos los Teléfonos"
                For k% = 0 To lstTelefonos.ListCount - 1
                    lstTelefonos.ListIndex = k%
                    pos1% = InStr(1, lstTelefonos.Text, Chr(9))
                    pos2% = InStr(pos1% + 1, lstTelefonos.Text, Chr(9))
                     'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "I"
                     'PMPasoValores SqlConn&, "@i_ente", 0, SQLINT4&, (pnlEnte.Caption)
                     'PMPasoValores SqlConn&, "@i_direccion", 0, SQLINT1&, (lblNumDireccion.Caption)
                     'PMPasoValores SqlConn&, "@i_valor", 0, SQLVARCHAR&, Mid$(lstTelefonos.Text, 1, pos1% - 1)
                     'PMPasoValores SqlConn&, "@i_tipo_telefono", 0, SQLCHAR&, Mid$(lstTelefonos.Text, pos1% + 1, pos2% - pos1% - 1)
                     'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "111"
                     'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_telefono", True, "Teléfonos") Then
                         'PMMapeaObjeto SqlConn&, lblNumTelefonos
                         'PMChequea SqlConn&
                    'End If
                Next k%
            Else
                 'PMChequea SqlConn&
            End If
            VLTranOk$ = "S"
        'Else
            'fraDireccion.Enabled = True
            'cmdTransmitir.Enabled = True
            'VLTranOk$ = "N"
            'Exit Sub
        'End If
        cmdBuscar_click

End Sub

Private Sub jmbIngresar_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Transmitir la Operación"
End Sub

Private Sub jmbLimpiar_Click()
        lblNumDireccion.Caption = ""
        txtDescripcionDireccion.Text = ""
        'PMPasoValores SqlConn&, "@i_operacion", 0, SQLCHAR&, "H"
        'PMPasoValores SqlConn&, "@i_tipo_h", 0, SQLCHAR&, "V"
        'PMPasoValores SqlConn&, "@i_ciudad", 0, SQLINT4&, (txtCiudad(0).Text)
        'PMPasoValores SqlConn&, "@i_provincia", 0, SQLINT4&, (txtProvincia_dir.Text)
        'PMPasoValores SqlConn&, "@i_parroquia", 0, SQLINT4&, (txtParroquia.Text)
        'codigo de transaccion
        'PMPasoValores SqlConn&, "@t_trn", 0, SQLINT2&, "1559"
         ReDim VTArreglo(4) As String
        'If FMTransmitirRPC(SqlConn&, ServerName$, "cobis", "sp_parroquia", True, "Ayuda Sector") Then
              'VTR% = FMMapeaArreglo(SqlConn&, VTArreglo())
              'PMChequea SqlConn&
        'End If
        txtTelefono.MaxLength = 12
        txtTelefono.Text = ""
        VLOperacion$ = "I"
        fraDireccion.Enabled = True
        fraDireccion.Caption = "Dirección [Creación]"
        jmbIngresar.Enabled = True
        VLPaso% = True
End Sub

Private Sub jmbLimpiar_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Limpiar los Campos de Ingreso de Datos"
End Sub

Private Sub jmbSalir_Click()
    Unload FApGeneral
End Sub

Private Sub jmbSalir_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Salir de la Forma"
End Sub

Private Sub txtDescripcionDireccion_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Dirección "
    txtDescripcionDireccion.SelStart = 0
    txtDescripcionDireccion.SelLength = Len(txtDescripcionDireccion.Text)
End Sub

Private Sub txtDescripcionDireccion_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea alfanumerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'21/Ene/94  M.Davila            Emisión Inicial
'*********************************************************
    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
End Sub

Private Sub txtDescripcionDireccion_LostFocus()
'*********************************************************
'Objetivo:  Forza el texto a mayúsculas y elimina los blan-
'           cos a izquierda y derecha.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    txtDescripcionDireccion.Text = Trim$(UCase$((txtDescripcionDireccion.Text)))
End Sub

Private Sub txtDireccion_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    
    FPrincipal!pnlHelpLine.Caption = " Dirección"
    txtDireccion.SelStart = 0
    txtDireccion.SelLength = Len(txtDireccion.Text)
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea alfanumerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'21/Ene/94  M.Davila            Emisión Inicial
'*********************************************************
    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
End Sub

Private Sub txtDireccion_lostfocus()
'*********************************************************
'Objetivo:  Forza el texto a mayúsculas y elimina los blan-
'           cos a izquierda y derecha.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    txtDireccion.Text = Trim$(UCase$((txtDireccion.Text)))
End Sub

Private Sub txtEmpresa_Change()
'*********************************************************
'Objetivo:  Determina que el campo empresa ha cambiado.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    VLPaso% = False
End Sub

Private Sub txtEmpresa_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado e ini-
'           cializa la variable que chequea los cambios
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Código de la Compañía en la que Trabaja  [F5] Ayuda"
    txtEmpresa.SelStart = 0
    txtEmpresa.SelLength = Len(txtEmpresa.Text)
End Sub

Private Sub txtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
'*********************************************************
'Objetivo:  Llama a la Forma de FBuscarCliente, la cual
'           muestra los códigos de empresa, cuando se
'           presiona la tecla F5
'Input   :  KeyCode     código de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    If KeyCode = F5% Then
        VLPaso% = True
        FBuscarCliente!optCliente(0).Enabled = False
'Modificación por cambios en la forma de búsqueda de Clientes
' por adaptación de clientes IFIs
        FBuscarCliente!optCliente(2).Enabled = True
        FBuscarCliente!optCliente(3).Enabled = True

        FBuscarCliente!optCliente(1).Value = True
        FBuscarCliente.Show 1
        FBuscarCliente!optCliente(0).Enabled = True
        If VGBusqueda(1) <> "" Then
            txtEmpresa.Text = VGBusqueda(1)
            lblEmpresa.Caption = VGBusqueda(2)
            VLPaso% = True
            SendKeys "{TAB}"
        Else
            VLPaso% = False
        End If
    End If
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea numerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'21/Ene/94  M.Davila            Emisión Inicial
'*********************************************************
    If (KeyAscii <> 8) And (KeyAscii% <> 32) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) Then
        KeyAscii% = 0
    End If
End Sub


Private Sub txtObservacionF_LostFocus()
'*********************************************************
'Objetivo:  elimina blancos a izquierda y derecha
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'02/May/94  M.Davila            Emisión Inicial
'*********************************************************
    txtObservacionF.Text = Trim$(txtObservacionF.Text)
End Sub

Private Sub txtObservacionP_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Observaciones a la Referencia"
    txtObservacionP.SelStart = 0
    txtObservacionP.SelLength = Len(txtObservacionP.Text)
End Sub

Private Sub txtObservacionP_KeyPress(KeyAscii As Integer)
    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
End Sub

Private Sub txtObservacionP_LostFocus()
    txtObservacionP.Text = Trim$(txtObservacionP.Text)
End Sub

Private Sub txtOficina_Change()
'*********************************************************
'Objetivo:  Determina que el campo oficina ha cambiado.
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'27/Abr/94  M.Davila            Emisión Inicial
'*********************************************************
    VLPaso% = False
End Sub

Private Sub txtOficina_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado e ini-
'           cializa la variable que chequea los cambios
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'27/Abr/94  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Oficina más cercana a la Dirección  [F5] Ayuda"
    txtOficina.SelStart = 0
    txtOficina.SelLength = Len(txtOficina.Text)
End Sub


Private Sub txtTelefono_GotFocus()
'*********************************************************
'Objetivo:  Despliega el mensaje de ayuda apropiado
'Input   :  ninguno
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Número Telefónico"
    txtTelefono.SelStart = 0
    'i% = Len(Trim$(txtTelefono.Text))
    'If i% < 12 Then
    '    For j% = 1 To 12 - i%
    '        txtTelefono.Text = txtTelefono.Text + " "
    '    Next j%
    'End If
    txtTelefono.SelLength = Len(txtTelefono.Text)
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  controla que el texto sea alfanumerico
'Input   :  KeyAscii    codigo ascii de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'21/Ene/94  M.Davila            Emisión Inicial
'*********************************************************
    KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
End Sub


