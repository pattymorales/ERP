VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FCompania 
   Caption         =   "Compañía"
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
      TabIndex        =   14
      Top             =   1920
      Width           =   7410
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   4
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
         Picture         =   "Fcompania.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
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
         Picture         =   "Fcompania.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   6
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
         Picture         =   "Fcompania.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
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
         Picture         =   "Fcompania.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
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
         Picture         =   "Fcompania.frx":1CE8
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
      TabIndex        =   9
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtDireccion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   65
         TabIndex        =   2
         Top             =   900
         Width           =   4995
      End
      Begin VB.TextBox txtRuc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   13
         TabIndex        =   0
         Top             =   300
         Width           =   1905
      End
      Begin VB.TextBox txtnombre 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   65
         TabIndex        =   1
         Top             =   600
         Width           =   4995
      End
      Begin VB.TextBox txtRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   65
         TabIndex        =   3
         Top             =   1200
         Width           =   4995
      End
      Begin VB.Label lblnombre 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Nombre o Razón Social:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   600
         Width           =   2070
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
         Left            =   105
         TabIndex        =   12
         Top             =   900
         Width           =   885
      End
      Begin VB.Label lblRuc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
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
         Left            =   105
         TabIndex        =   11
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblRepresentante 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Representante Legal:"
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
         Left            =   105
         TabIndex        =   10
         Top             =   1200
         Width           =   1845
      End
   End
End
Attribute VB_Name = "FCompania"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PLIngresar()
Dim rstcompania As Recordset
Dim errObj As Error
Dim VTCodigo As String

'Validacion de datos
    If txtRuc.Text = "" Then
        MsgBox "Debe ingresar el RUC de la compania", vbCritical, "Warning"
        txtRuc.SetFocus
        Exit Sub
    End If
    If Len(txtRuc.Text) <> 0 And Len(txtRuc.Text) <> 13 Then
        MsgBox "El número de caracteres del RUC debe ser de 13", vbCritical, "Ingresar"
        txtRuc.SetFocus
        Exit Sub
    End If
    If txtnombre.Text = "" Then
        MsgBox "Debe ingresar el nombre de la compania", vbCritical, "Warning"
        txtnombre.SetFocus
        Exit Sub
    End If
    If txtDireccion.Text = "" Then
        MsgBox "Debe ingresar la direccion", vbCritical, "Warning"
        txtDireccion.SetFocus
        Exit Sub
    End If
    If txtRepresentante.Text = "" Then
        MsgBox "Debe ingresar el nombre del representante", vbCritical, "Warning"
        txtRepresentante.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    strT = "SELECT co_ruc FROM er_compania WHERE co_ruc = '" & Trim$(txtRuc.Text) & "'"
    Set rstcompania = dbErp.OpenRecordset(strT)
    If Not rstcompania.EOF Then
        MsgBox "La compania ya existe, no se puede ingresar", vbCritical, "Ingresar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_compania @i_operacion = 'I', @i_ruc = '" & txtRuc.Text & _
    "', @i_nombre = '" & txtnombre.Text & "', @i_direccion = '" & _
    txtDireccion.Text & "', @i_representante = '" & txtRepresentante.Text & "'"
    Set rstcompania = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Compania creada correctamente", vbOKOnly, "Ingreso"
    txtRuc.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
ErrorInsertar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Insertar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al insertar la compania", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbbuscar_Click()
    VGForma = "Fcompania"
    Fbuscacompania.Show 1
    If txtRuc.Text <> "" Then
        txtRuc.Enabled = False
    Else
        txtRuc.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    PLBuscar_compania
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
Dim rstcompania As Recordset
Dim errObj As Error

'Validacion de datos
    If txtRuc.Text = "" Then
        MsgBox "Debe ingresar el RUC de la compania", vbCritical, "Warning"
        Exit Sub
    End If
     If txtnombre.Text = "" Then
        MsgBox "Debe ingresar el nombre de la compania", vbCritical, "Warning"
        txtnombre.SetFocus
        Exit Sub
    End If
    If txtDireccion.Text = "" Then
        MsgBox "Debe ingresar la dirección de la compania", vbCritical, "Warning"
        txtDireccion.SetFocus
        Exit Sub
    End If
    If txtRepresentante.Text = "" Then
        MsgBox "Debe ingresar el nombre del representante", vbCritical, "Warning"
        txtRepresentante.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorModificar
Screen.MousePointer = 11
    strT = "SELECT co_ruc FROM er_compania WHERE co_ruc = '" & Trim$(txtRuc.Text) & "'"
    Set rstcompania = dbErp.OpenRecordset(strT)
    If rstcompania.EOF Then
        MsgBox "La compania no existe, no se puede modificar", vbCritical, "Modificar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_compania @i_operacion = 'U', @i_ruc = '" & _
    txtRuc.Text & "', @i_nombre = '" & txtnombre.Text & "'," & _
    " @i_direccion  = '" & txtDireccion.Text & "', @i_representante = '" & txtRepresentante.Text & "'"
    Set rstcompania = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Compañía modificada correctamente", vbOKOnly, "Modificar"
    Screen.MousePointer = 0
    Exit Sub
ErrorModificar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Modificar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al modificar la compania", vbCritical, "Modificar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstcompania As Recordset
Dim errObj As Error
        
    If txtRuc.Text = "" Then
        MsgBox "Debe ingresar el RUC de la compania", vbCritical, "Warning"
        Exit Sub
    End If
    VT% = MsgBox("Esta seguro de eliminar la compania:  " & txtRuc.Text & " ?", vbYesNo, "Eliminar")
    If VT% = 7 Then
        Exit Sub
    End If
On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    strT = "SELECT co_ruc FROM er_compania WHERE co_ruc = '" & Trim$(txtRuc.Text) & "'"
    Set rstcompania = dbErp.OpenRecordset(strT)
    If rstcompania.EOF Then
        MsgBox "La compania no existe", vbCritical, "Eliminar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_compania @i_operacion = 'D', @i_ruc = '" & txtRuc.Text & "'"
    Set rstcompania = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "La compania ha sido eliminada correctamente", vbOKOnly, "Eliminar"
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
        Var% = MsgBox("Error al eliminar la compania", vbCritical, "Eliminar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLLimpiar()
    txtRuc.Text = ""
    txtnombre.Text = ""
    txtDireccion.Text = ""
    txtRepresentante.Text = ""
    txtRuc.Enabled = True
End Sub

Private Sub PLBuscar_compania()
Dim rstcompania As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    Command_SQLx = "sp_compania @i_operacion ='H', @i_tipo = 'A', @i_modo = 0"
    Set rstcompania = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    If Not rstcompania.EOF Then
        txtRuc.Text = rstcompania(0)
        txtnombre.Text = rstcompania(1)
        txtDireccion.Text = rstcompania(2)
        txtRepresentante.Text = rstcompania(3)
        txtRuc.Enabled = False
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar la compañía", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub


Private Sub txtRuc_LostFocus()
    If Len(txtRuc.Text) <> 0 And Len(txtRuc.Text) <> 13 Then
        MsgBox "El número de caracteres del RUC debe ser de 13", vbCritical, "Ingresar"
        txtRuc.SetFocus
    End If
End Sub
