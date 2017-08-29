VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FApCliente 
   Appearance      =   0  'Flat
   Caption         =   "Datos Generales de Clientes"
   ClientHeight    =   5490
   ClientLeft      =   975
   ClientTop       =   2025
   ClientWidth     =   9525
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
   HelpContextID   =   1015
   Icon            =   "FAPPER.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   9525
   Tag             =   "ApCliente"
   WindowState     =   2  'Maximized
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   50
      TabIndex        =   29
      Top             =   3360
      Width           =   7410
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   12
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
         Picture         =   "FAPPER.frx":030A
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   13
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
         Picture         =   "FAPPER.frx":08A4
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   14
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
         Picture         =   "FAPPER.frx":117E
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   15
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
         Picture         =   "FAPPER.frx":1A58
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
         TabIndex        =   16
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
         Picture         =   "FAPPER.frx":1FF2
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos Generales"
      ForeColor       =   &H00000080&
      Height          =   3255
      Left            =   50
      TabIndex        =   18
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtemail 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         MaxLength       =   64
         TabIndex        =   11
         Top             =   2700
         Width           =   5265
      End
      Begin Threed.SSOption ssoCedula 
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Cédula"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin VB.ComboBox cboSexo 
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         MaxLength       =   64
         TabIndex        =   5
         Top             =   1200
         Width           =   5265
      End
      Begin VB.TextBox txtMaterno 
         Height          =   285
         Left            =   2160
         MaxLength       =   32
         TabIndex        =   7
         Top             =   1800
         Width           =   3360
      End
      Begin VB.TextBox txtPaterno 
         Height          =   285
         Left            =   2160
         MaxLength       =   32
         TabIndex        =   6
         Top             =   1500
         Width           =   3360
      End
      Begin MSMask.MaskEdBox txtDocumentoID 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   900
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskNacimiento 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   2100
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskIngreso 
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   2040
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption ssoRuc 
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "RUC"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption ssoPasaporte 
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   600
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Pasaporte"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblEmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblTipo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblNumero 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2160
         TabIndex        =   0
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lblFechaIngreso 
         Appearance      =   0  'Flat
         Caption         =   "Fecha de Ingreso:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4230
         TabIndex        =   26
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label lblFechaNac 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label lblNumeroDoc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Número de Documento:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   900
         Width           =   2010
      End
      Begin VB.Label lblApellidoM 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lblApellidoP 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1500
         Width           =   1470
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lblIdentificacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "No. de Identificación:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1860
      End
      Begin VB.Label lblSexo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sexo:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   495
      End
   End
   Begin JMButton.JMBcontrol jmbDireccion 
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "FAPPER.frx":28CC
      Caption         =   "Dirección"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   7545
      X2              =   7545
      Y1              =   5610
      Y2              =   15
   End
End
Attribute VB_Name = "FApCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Constantes
Const F5% = 116     'codigo de la tecla F5

'Variables Locales
Dim VLPaso%         'controla el cambio de los campos
Dim VLOficina$      'variable local con codigo de oficina
Dim VLFilial$       'variable local con codigo de filial
Dim VLPrimeraVez%   'para inicializar la forma para operación
Dim VLFormato$      'guarda el formato de fecha
Dim VLError%        'toma valor de uno cuando hay error
'Constantes Locales
Const IDCANCEL = 2  ' Cancel button pressed
Const IDYES = 6     ' Yes button pressed
Const IDNO = 7      ' No button pressed

Private Sub Form_Activate()
    FApCliente.HelpContextID = 101
    cmdLimpiar.Enabled = True
    FApCliente.Caption = "Datos Generales Clientes"
    MskNacimiento.Text = "01/01/1900"
    VLPaso% = True
    cboSexo.AddItem "Femenino"
    cboSexo.AddItem "Masculino"
End Sub

Private Sub Form_Load()
    VLPrimeraVez% = True
    'carga de formato de fecha
    MskNacimiento.Mask = FMMascaraFecha("mm/dd/yyyy")
    jmbIngresar.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FPrincipal.pnlHelpLine.Caption = ""
    FPrincipal.pnlTransaccionLine.Caption = ""
End Sub

Private Sub jmbDireccion_Click()
    FApGeneral.Show 0
End Sub

Private Sub jmbEliminar_Click()
    PLEliminar
End Sub

Private Sub jmbIngresar_Click()
    PLIngresar
End Sub

Private Sub PLIngresar()
Dim rstCliente As Recordset
Dim errObj As Error
Dim VTCodigo As String
    'chequeo de campos mandatorios
    If txtNombre.Text = "" Then
        MsgBox "El Campo Nombre es Mandatorio", 48, "Control Ingreso de Datos"
        txtNombre.SetFocus
        Exit Sub
    End If
    If txtPaterno.Text = "" Then
        MsgBox "El Campo Apellido Paterno es Mandatorio", 48, "Control Ingreso de Datos"
        txtPaterno.SetFocus
        Exit Sub
    End If
    'Calcular por diferencia de fecha la edad de la persona
    VTNow$ = Format$(Now, VLFormato$)
    VTanio% = FMDateDiff("y", (MskNacimiento.FormattedText), VTNow$, VGFormatoFecha)
    'Si es mayor de edad cédula o pasaporte es campo mandatorio
    If (VTanio% >= VGMDE%) Or (MskNacimiento.ClipText = "") Then
    'RCS 15/nov/2007 Controlar que cuando sea proveedor internacional no sea obligatorio
    If (txtDocumentoID.Text = "__________" Or txtDocumentoID.Text = "") Then
            MsgBox "El Campo Documento es Mandatorio", 48, "Control Ingreso de Datos"
            txtDocumentoID.SetFocus
            Exit Sub
        End If
    End If
    If cboSexo.Text = "" Then
        cboSexo.ListIndex = 2
    End If

On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    If ssoCedula.Value Then
        VTTipo_doc$ = "C"
    Else
        If ssoRuc.Value Then
            VTTipo_doc$ = "R"
        Else
            If ssoPasaporte.Value Then
                VTTipo_doc$ = "P"
            End If
        End If
    End If
    Command_SQLx = "sp_ente @i_operacion = 'I', @i_tipo_doc = '" & VTTipo_doc & "'," & _
    " @i_documento& '" & txtDocumentoID.FormattedText & " ', @i_nombre = '" & txtNombre.Text & "', " & _
    " @i_p_apellido = ' " & txtPaterno.Text & "', @i_s_apellido = '" & txtMaterno.Text & " ', " & _
    " @i_fecha_nac = '" & FMConvFecha((MskNacimiento.FormattedText), VGFormatoFecha, "mm/dd/yyyy") & "', " & _
    " @i_fecha_ing = '" & FMConvFecha((MskIngreso.FormattedText), VGFormatoFecha, "mm/dd/yyyy") & "', " & _
    " @i_sexo = '" & cboSexo.Text & "', @i_email = '" & txtemail.Text & "'"
    Set rstCliente = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VTCodigo = rstCliente(0)
    lblNumero.Caption = VTCodigo
    MsgBox "Cliente creado correctamente", vbOKOnly, "Ingreso"
    txtNombre.Enabled = False
    txtPaterno.Enabled = False
    txtMaterno.Enabled = False
    txtDocumentoID.Enabled = False
    MskNacimiento.Enabled = False
    cboSexo.Enabled = False
    jmbIngresar.Enabled = False
    jmbDireccion.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrorInsertar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Insertar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al insertar el Cliente", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbIngresar_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Grabar la información"
End Sub

Private Sub jmbLimpiar_Click()
'Limpiar
    PLLimpiar
End Sub

Private Sub PLLimpiar()
ReDim VTArreglo(5) As String
    lblNumero.Caption = ""
    txtDocumentoID.Text = ""
    txtNombre.Text = ""
    txtPaterno.Text = ""
    txtMaterno.Text = ""
    txtDocumentoID.Text = "__________"
    txtNacimiento.Mask = ""
    txtNacimiento.Text = "01/01/1900"
    txtNacimiento.Mask = FMMascaraFecha(VLFormato$)
    cboSexo.Text = ""
    txtNombre.Enabled = True
    txtPaterno.Enabled = True
    txtMaterno.Enabled = True
    txtDocumentoID.Enabled = True
    txtNacimiento.Enabled = True
    MskIngreso.Enabled = False
    jmbModificar.Enabled = True
    txtNombre.SetFocus
    VLPaso% = True
End Sub

Private Sub jmbLimpiar_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Limpiar los Campos de Ingreso de Datos"
End Sub

Private Sub jmbModificar_Click()
    PLModificar
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub jmbSalir_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Salir de la Forma"
End Sub

Private Sub ssoCedula_Click(Value As Integer)
    If ssoCedula.Value = True Then
        txtDocumentoID.Mask = "##########"
    End If
End Sub

Private Sub ssoPasaporte_Click(Value As Integer)
     If ssoPasaporte.Value = True Then
        txtDocumentoID.Mask = ""
    End If
End Sub

Private Sub ssoRuc_Click(Value As Integer)
    If ssoRuc.Value = True Then
        txtDocumentoID.Mask = "#############"
    End If
End Sub

Private Sub txtdocumentoid_Change()
    VLPaso% = False
End Sub

Private Sub txtdocumentoid_GotFocus()
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Número de Documento de Indentidad"
    FPrincipal!pnlTransaccionLine.Caption = ""
    txtDocumentoID.SelStart = 0
    txtDocumentoID.SelLength = Len(txtDocumentoID.Text)
End Sub

Private Sub txtdocumentoid_KeyPress(KeyAscii As Integer)
    If ssoRuc.Value = True Then
        If (KeyAscii%) = 32 Then
            KeyAscii% = 0
        Else
           If (KeyAscii <> 8) And (KeyAscii% <> 32) And ((KeyAscii% < 48) Or (KeyAscii% > 57)) Then
              KeyAscii% = 0
           End If
        End If
    End If
End Sub

Private Sub txtdocumentoid_LostFocus()
'*********************************************************
'Objetivo:  Calcula que el numero de cedula sea valido,
'           chequea el digito verificador
'*********************************************************
    If ssoCedula.Value = True Then
        Dim VTDigito%
        If Not VLPaso% Then
            If txtDocumentoID.Text <> "__________" Then
                VTDigito% = 0
                If InStr(1, txtDocumentoID.Text, "_") <> 0 Then
                    MsgBox "Número de Cédula Incorrecto", 48, "Control Ingreso de Datos"
                    txtDocumentoID.SetFocus
                    Exit Sub
                End If
                For i% = 1 To 9
                    If i% Mod 2 <> 0 Then
                        If Mid$(txtDocumentoID.Text, i%, 1) <> "9" Then
                            VTDigito% = VTDigito% + ((Val(Mid$(txtDocumentoID.Text, i%, 1)) * 2) Mod 9)
                        Else
                            VTDigito% = VTDigito% + 9
                        End If
                    Else
                        VTDigito% = VTDigito% + Val(Mid$(txtDocumentoID.Text, i%, 1))
                    End If
                Next i%
                VTDigito% = 10 - (VTDigito% Mod 10)
                If VTDigito% = 10 Then
                    VTDigito% = 0
                End If
                VTVerificador% = Val(Mid$(txtDocumentoID.Text, 10, 1))
                If VTVerificador% <> VTDigito% Then
                    MsgBox "Dígito Verificador NO corresponde...", 0, "Control de Ingreso de Datos"
                    'SendKeys "{TAB}"
                End If
            End If
        End If
    End If
    If ssoRuc.Value = True Then
        If Len(txtRuc.Text) <> 13 And Len(txtRuc.Text) <> 0 Then
            MsgBox "Ruc debe ser de 13 digitos", 48, "Control Ingreso de Datos"
            txtRuc.SetFocus
            Exit Sub
        End If
    End If
    txtDocumentoID.Text = Trim$(txtDocumentoID.Text)
End Sub

Private Sub mskingreso_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Fecha de Ingreso al Pais  [" & VLFormato$ & "]"
    MskIngreso.SelStart = 0
    MskIngreso.SelLength = Len(MskIngreso.Text)
End Sub

Private Sub mskingreso_LostFocus()
Dim VT%
Dim VTNow$
    If MskIngreso.ClipText <> "" Then
        VT% = FMVerFormato((MskIngreso.FormattedText), VLFormato$)
        If Not VT% Then
            MsgBox "Formato de Fecha Inválido", 48, "Control Ingreso de Datos"
            VLError% = True
            MskIngreso.SetFocus
            Exit Sub
        Else
            'chequear que la fecha sea mayor a la fecha de nacimiento
            If txtNacimiento.ClipText <> "" Then
                If FMDateDiff("d", (MskIngreso.FormattedText), (txtNacimiento.FormattedText), VLFormato$) > 0 Then
                    MsgBox "Fecha debe ser mayor a la Fecha de Nacimiento", 48, "Control Ingreso de Datos"
                    VLError% = True
                    MskIngreso.SetFocus
                    Exit Sub
                End If
            End If
            'chequear que la fecha sea menor o igual que la del sistema
            VTNow$ = Format$(Now, VLFormato$)
            If FMDateDiff("d", (MskIngreso.FormattedText), VTNow$, VLFormato$) < 0 Then
                MsgBox "Fecha fuera de Rango", 48, "Control Ingreso de Datos"
                VLError% = True
                MskIngreso.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtMaterno_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Apellido Materno del Cliente"
    FPrincipal!pnlTransaccionLine.Caption = ""
    txtMaterno.SelStart = 0
    txtMaterno.SelLength = Len(txtMaterno.Text)
End Sub

Private Sub txtMaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 39 Or KeyAscii% = 209 Or KeyAscii% = 241 Then
        KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
    Else
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If
End Sub

Private Sub txtMaterno_LostFocus()
    txtMaterno.Text = Trim$(UCase$((txtMaterno.Text)))
End Sub

Private Sub mskNacimiento_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    VLPaso% = True
    FPrincipal!pnlHelpLine.Caption = " Fecha de Nacimiento del Cliente  [" & VLFormato$ & "]"
    txtNacimiento.SelStart = 0
    txtNacimiento.SelLength = Len(txtNacimiento.Text)
End Sub

Private Sub mskNacimiento_LostFocus()
Dim VT%
Dim VTNow$
    If MskNacimiento.Text <> "" Then
        VT% = FMVerFormato((txtNacimiento.FormattedText), VLFormato$)
        If Not VT% Then
            MsgBox "Formato de Fecha Inválido", 48, "Control Ingreso de Datos"
            VLError% = True
            MskNacimiento.SetFocus
            Exit Sub
        Else
            'chequear que la fecha sea menor que la del sistema
            VTNow$ = Format$(Now, VLFormato$)
            If FMDateDiff("d", (MskNacimiento.Text), VTNow$, VLFormato$) <= 0 Then
                MsgBox "Fecha fuera de Rango", 48, "Control Ingreso de Datos"
                VLError% = True
                MskNacimiento.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtNombre_GotFocus()
    FPrincipal!pnlHelpLine.Caption = " Nombres completos del Cliente"
    FPrincipal!pnlTransaccionLine.Caption = ""
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 39 Or KeyAscii% = 209 Or KeyAscii% = 241 Then
        KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
    Else
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If
End Sub

Private Sub txtNombre_LostFocus()
    txtNombre.Text = Trim$(UCase$((txtNombre.Text)))
End Sub

Private Sub txtPaterno_GotFocus()
    FPrincipal!pnlTransaccionLine.Caption = ""
    FPrincipal!pnlHelpLine.Caption = " Apellido Paterno del Cliente"
    txtPaterno.SelStart = 0
    txtPaterno.SelLength = Len(txtPaterno.Text)
End Sub

Private Sub txtPaterno_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 39 Or KeyAscii% = 209 Or KeyAscii% = 241 Then
       KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
    Else
        If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% < 65 Or KeyAscii% > 90) And (KeyAscii% < 97 Or KeyAscii% > 122) Then
            KeyAscii% = 0
        Else
            KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
        End If
    End If
End Sub

Private Sub txtPaterno_LostFocus()
    txtPaterno.Text = Trim$(UCase$((txtPaterno.Text)))
End Sub

Private Sub PLModificar()
Dim rstCliente As Recordset
Dim errObj As Error
Dim VTCodigo As String
    'chequeo de campos mandatorios
    If lblNumero.Caption = "" Then
        MsgBox "El Número de Identificación es obligatorio", 48, "Control Ingreso de Datos"
        Exit Sub
    End If
    If txtNombre.Text = "" Then
        MsgBox "El Campo Nombre es Mandatorio", 48, "Control Ingreso de Datos"
        txtNombre.SetFocus
        Exit Sub
    End If
    If txtPaterno.Text = "" Then
        MsgBox "El Campo Apellido Paterno es Mandatorio", 48, "Control Ingreso de Datos"
        txtPaterno.SetFocus
        Exit Sub
    End If
    'Calcular por diferencia de fecha la edad de la persona
    VTNow$ = Format$(Now, VLFormato$)
    VTanio% = FMDateDiff("y", (MskNacimiento.FormattedText), VTNow$, VGFormatoFecha)
    'Si es mayor de edad cédula o pasaporte es campo mandatorio
    If (VTanio% >= VGMDE%) Or (MskNacimiento.ClipText = "") Then
    'RCS 15/nov/2007 Controlar que cuando sea proveedor internacional no sea obligatorio
    If (txtDocumentoID.Text = "__________" Or txtDocumentoID.Text = "") Then
            MsgBox "El Campo Documento es Mandatorio", 48, "Control Ingreso de Datos"
            txtDocumentoID.SetFocus
            Exit Sub
        End If
    End If
    If cboSexo.Text = "" Then
        cboSexo.ListIndex = 2
    End If

On Error GoTo ErrorModificar
    Screen.MousePointer = 11
    If ssoCedula.Value Then
        VTTipo_doc$ = "C"
    Else
        If ssoRuc.Value Then
            VTTipo_doc$ = "R"
        Else
            If ssoPasaporte.Value Then
                VTTipo_doc$ = "P"
            End If
        End If
    End If
    Command_SQLx = "sp_ente @i_operacion = 'U', @i_tipo_doc = '" & VTTipo_doc & "'," & _
    " @i_documento& '" & txtDocumentoID.FormattedText & " ', @i_nombre = '" & txtNombre.Text & "', " & _
    " @i_p_apellido = ' " & txtPaterno.Text & "', @i_s_apellido = '" & txtMaterno.Text & " ', " & _
    " @i_fecha_nac = '" & FMConvFecha((MskNacimiento.FormattedText), VGFormatoFecha, "mm/dd/yyyy") & "', " & _
    " @i_fecha_ing = '" & FMConvFecha((MskIngreso.FormattedText), VGFormatoFecha, "mm/dd/yyyy") & "', " & _
    " @i_sexo = '" & cboSexo.Text & "', @i_email = '" & txtemail.Text & "'"
    Set rstCliente = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VTCodigo = rstCliente(0)
    lblNumero.Caption = VTCodigo
    MsgBox "Cliente modificado correctamente", vbOKOnly, "Ingreso"
    jmbDireccion.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrorModificar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Modificar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al modificar el Cliente", vbCritical, "Modificar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstCliente As Recordset
Dim errObj As Error
Dim VTCodigo As String
    'chequeo de campos mandatorios
    If lblNumero.Caption = "" Then
        MsgBox "El Número de Identificación es obligatorio", 48, "Control Ingreso de Datos"
        Exit Sub
    End If
    
On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    If ssoCedula.Value Then
        VTTipo_doc$ = "C"
    Else
        If ssoRuc.Value Then
            VTTipo_doc$ = "R"
        Else
            If ssoPasaporte.Value Then
                VTTipo_doc$ = "P"
            End If
        End If
    End If
    Command_SQLx = "sp_ente @i_operacion = 'D', @i_ente = " & lblNumero.Caption
    Set rstCliente = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Cliente eliminado correctamente", vbOKOnly, "Eliminación"
    PLLimpiar
    Screen.MousePointer = 0
    Exit Sub
ErrorEliminar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Eliminar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al eliminar el Cliente", vbCritical, "Eliminar"
    End If
    Screen.MousePointer = 0
End Sub
