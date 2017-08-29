VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FSucursal 
   Caption         =   "Sucursal"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   7575
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
         Picture         =   "Fsucursal.frx":0000
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
         Picture         =   "Fsucursal.frx":059A
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
         Picture         =   "Fsucursal.frx":0E74
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
         Picture         =   "Fsucursal.frx":174E
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
         Picture         =   "Fsucursal.frx":1CE8
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtPtofacturacion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   64
         TabIndex        =   10
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtPtoemision 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   64
         TabIndex        =   9
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   64
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox txtDireccion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   64
         TabIndex        =   2
         Top             =   900
         Width           =   5355
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   64
         TabIndex        =   1
         Top             =   600
         Width           =   5355
      End
      Begin VB.Label lblPtofacturacion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Punto de Facturación:"
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
         Top             =   1500
         Width           =   1905
      End
      Begin VB.Label lblPtoemision 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Punto de Emisión:"
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
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código de la Sucursal:"
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
         Width           =   1935
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         TabIndex        =   4
         Top             =   600
         Width           =   1080
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
         Left            =   135
         TabIndex        =   3
         Top             =   900
         Width           =   885
      End
   End
   Begin JMButton.JMBcontrol jmbBuscar 
      Height          =   495
      Left            =   7800
      TabIndex        =   17
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
      Picture         =   "Fsucursal.frx":25C2
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PLIngresar()
Dim rstSucursal As Recordset
Dim errObj As Error
Dim VTCodigo As String

'Validacion de datos
    If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción de la sucursal", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If txtDireccion.Text = "" Then
        MsgBox "Debe ingresar la dirección de la sucursal", vbCritical, "Warning"
        txtDireccion.SetFocus
        Exit Sub
    End If
    If txtPtoemision.Text = "" Then
        MsgBox "Debe ingresar el punto de emisión", vbCritical, "Warning"
        txtPtoemision.SetFocus
        Exit Sub
    End If
    If txtPtofacturacion.Text = "" Then
        MsgBox "Debe ingresar el punto de Facturación", vbCritical, "Warning"
        txtPtofacturacion.SetFocus
        Exit Sub
    End If

On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    Command_SQLx = "sp_sucursal @i_operacion = 'I', @i_descripcion = " & _
    "'" & txtDescripcion.Text & "', @i_direccion  = '" & txtDireccion.Text & "', @i_ptoemision = '" & txtPtoemision.Text & _
    "', @i_ptofacturacion = '" & txtPtofacturacion.Text & "'"
    Set rstSucursal = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VTCodigo = rstSucursal(0)
    txtCodigo.Text = VTCodigo
    MsgBox "Sucursal creada correctamente", vbOKOnly, "Ingreso"
    jmbIngresar.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
ErrorInsertar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Insertar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al insertar la sucursal", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbbuscar_Click()
    FbuscarSucursal.Show 1
    If txtCodigo.Text <> "" Then
        VLPaso% = False
        jmbIngresar.Enabled = False
    Else
        VLPaso% = False
    End If
End Sub

Private Sub jmbEliminar_Click()
    PLEliminar
End Sub

Private Sub jmbIngresar_Click()
    PLIngresar
End Sub

Private Sub jmbLimpiar_Click()
   PLLimpiar
End Sub

Private Sub jmbModificar_Click()
    PLModificar
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub PLModificar()
Dim rstSucursal As Recordset
Dim errObj As Error

'Validacion de datos
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código de la sucursal", vbCritical, "Warning"
        Exit Sub
    End If
     If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción de la sucursal", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If txtDireccion.Text = "" Then
        MsgBox "Debe ingresar la dirección de la sucursal", vbCritical, "Warning"
        txtDireccion.SetFocus
        Exit Sub
    End If
    If txtPtoemision.Text = "" Then
        MsgBox "Debe ingresar el punto de emisión", vbCritical, "Warning"
        txtPtoemision.SetFocus
        Exit Sub
    End If
    If txtPtofacturacion.Text = "" Then
        MsgBox "Debe ingresar el punto de Facturación", vbCritical, "Warning"
        txtPtofacturacion.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorModificar
Screen.MousePointer = 11
        Command_SQLx = "sp_sucursal @i_operacion = 'U', @i_codigo = '" & _
        txtCodigo.Text & "', @i_descripcion = '" & txtDescripcion.Text & "'," & _
        " @i_direccion  = '" & txtDireccion.Text & "', @i_ptoemision = '" & txtPtoemision.Text & _
        "', @i_ptofacturacion = '" & txtPtofacturacion.Text & "'"
        Set rstSucursal = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
        MsgBox "Sucursal modificada correctamente", vbOKOnly, "Modificar"
        Screen.MousePointer = 0
        jmbIngresar.Enabled = False
        Exit Sub
ErrorModificar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Modificar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al modificar la sucursal", vbCritical, "Modificar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstSucursal As Recordset
Dim errObj As Error
        
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código de la sucursal", vbCritical, "Warning"
        Exit Sub
    End If
    VT% = MsgBox("Esta seguro de eliminar la sucursal:  " & txtCodigo.Text & " ?", vbYesNo, "Eliminar")
    If VT% = 7 Then
        Exit Sub
    End If
On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    strT = "SELECT su_codigo FROM er_sucursal WHERE su_codigo = " & Trim$(txtCodigo.Text)
    Set rstSucursal = dbErp.OpenRecordset(strT)
    If rstSucursal.EOF Then
        MsgBox "La sucursal no existe", vbCritical, "Eliminar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_sucursal @i_operacion = 'D', @i_codigo = '" & txtCodigo.Text & "'"
    Set rstSucursal = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "La sucursal ha sido eliminada correctamente", vbOKOnly, "Eliminar"
    Screen.MousePointer = 0
    PLLimpiar
    Exit Sub
ErrorEliminar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Eliminar")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al eliminar la sucursal", vbCritical, "Eliminar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLLimpiar()
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    txtDireccion.Text = ""
    txtPtoemision.Text = ""
    txtPtofacturacion.Text = ""
    jmbIngresar.Enabled = True
End Sub
