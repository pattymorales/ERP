VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FUnidades 
   Caption         =   "Unidades de Stock"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   7455
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   3
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
         Picture         =   "FUnidades.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   4
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
         Picture         =   "FUnidades.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   5
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
         Picture         =   "FUnidades.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   6
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
         Picture         =   "FUnidades.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
         TabIndex        =   7
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
         Picture         =   "FUnidades.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos de la Unidad"
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
      TabIndex        =   9
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   64
         TabIndex        =   1
         Top             =   660
         Width           =   5355
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   64
         TabIndex        =   0
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblcUnidad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código de Unidad:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1530
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
         Left            =   135
         TabIndex        =   10
         Top             =   660
         Width           =   1080
      End
   End
   Begin JMButton.JMBcontrol jmbBuscar 
      Height          =   495
      Left            =   7680
      TabIndex        =   8
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
      Picture         =   "FUnidades.frx":25C2
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PLIngresar()
Dim rstUnidades As Recordset
Dim errObj As Error
Dim VTCodigo As String

'Validacion de datos
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código de la Unidad", vbCritical, "Warning"
        txtCodigo.SetFocus
        Exit Sub
    End If
    If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción de la Unidad", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    strT = "SELECT un_codigo FROM er_unidades WHERE un_codigo = '" & Trim$(txtCodigo.Text) & "'"
    Set rstUnidades = dbErp.OpenRecordset(strT)
    If Not rstUnidades.EOF Then
        MsgBox "El código de Unidad ya existe, no se puede ingresar", vbCritical, "Ingresar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_unidades @i_operacion = 'I', @i_codigo = '" & txtCodigo.Text & _
    "', @i_descripcion = '" & txtDescripcion.Text & "'"
    Set rstUnidades = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "unidad creada correctamente", vbOKOnly, "Ingreso"
    txtCodigo.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
ErrorInsertar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Insertar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al insertar la unidad", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbbuscar_Click()
    StoredProcedure = "sp_unidades"
    Fgrid_catalogo.Show 1
    If Temporales(1) <> "" Then
        txtCodigo.Text = Temporales(1)
        txtDescripcion.Text = Temporales(2)
        txtCodigo.Enabled = False
        jmbIngresar.Enabled = False
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
Dim rstUnidades As Recordset
Dim errObj As Error

'Validacion de datos
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código de la unidad", vbCritical, "Warning"
        Exit Sub
    End If
     If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción de la unidad", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorModificar
Screen.MousePointer = 11
    strT = "SELECT un_codigo FROM er_unidades WHERE un_codigo = '" & Trim$(txtCodigo.Text) & "'"
    Set rstUnidades = dbErp.OpenRecordset(strT)
    If rstUnidades.EOF Then
        MsgBox "La unidad no existe, no se puede modificar", vbCritical, "Modificar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_unidades @i_operacion = 'U', @i_codigo = '" & _
    txtCodigo.Text & "', @i_descripcion = '" & txtDescripcion.Text & "'"
    Set rstUnidades = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Unidad modificada correctamente", vbOKOnly, "Modificar"
    Screen.MousePointer = 0
    Exit Sub
ErrorModificar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Modificar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al modificar la unidad", vbCritical, "Modificar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstUnidades As Recordset
Dim errObj As Error
        
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código de la unidad", vbCritical, "Warning"
        Exit Sub
    End If
    VT% = MsgBox("Esta seguro de eliminar la unidad:  " & txtCodigo.Text & " ?", vbYesNo, "Eliminar")
    If VT% = 7 Then
        Exit Sub
    End If
On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    strT = "SELECT un_codigo FROM er_unidades WHERE un_codigo = '" & Trim$(txtCodigo.Text) & "'"
    Set rstUnidades = dbErp.OpenRecordset(strT)
    If rstUnidades.EOF Then
        MsgBox "La unidad no existe", vbCritical, "Eliminar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_unidades @i_operacion = 'D', @i_codigo = '" & txtCodigo.Text & "'"
    Set rstUnidades = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "La unidad ha sido eliminado correctamente", vbOKOnly, "Eliminar"
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
        Var% = MsgBox("Error al eliminar la unidad", vbCritical, "Eliminar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLLimpiar()
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    txtCodigo.Enabled = True
    jmbIngresar.Enabled = True
End Sub
