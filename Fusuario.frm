VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form Fusuario 
   Caption         =   "Usuario"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   7455
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   11
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
         Picture         =   "Fusuario.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
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
         Picture         =   "Fusuario.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
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
         Picture         =   "Fusuario.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
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
         Picture         =   "Fusuario.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
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
         Picture         =   "Fusuario.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmPanel 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtDato 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   1500
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2040
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1500
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "Usuario Activo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Tag             =   "30002"
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2040
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Tag             =   "10048"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de acceso:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Tag             =   "10045"
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repita la clave:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Tag             =   "10045"
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Tag             =   "10049"
         Top             =   360
         Width           =   675
      End
   End
   Begin JMButton.JMBcontrol jmbbuscar 
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "Fusuario.frx":25C2
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin VB.Line lnllinea 
      BorderColor     =   &H00808080&
      X1              =   7800
      X2              =   7800
      Y1              =   0
      Y2              =   5760
   End
End
Attribute VB_Name = "Fusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub PLIngreso()
Dim rstLogin As Recordset
Dim errObj As Error
Dim l As Long
Dim Command_SQLx As String
    
If txtDato(1).Text = "" Then
    MsgBox "Debe ingresar el login del usuario", vbCritical, "Warning"
    txtDato(1).SetFocus
    Exit Sub
End If
If txtDato(2).Text = "" Then
    MsgBox "El password del usuario es obligatorio", vbCritical
    txtDato(2).SetFocus
    Exit Sub
End If
'Verifica las Password
If txtDato(2).Text <> txtDato(3).Text Then
    MsgBox "El password del usuario no está correcto", vbCritical
    txtDato(2).SetFocus
    Exit Sub
End If
On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    strT = "SELECT us_codigo FROM er_user WHERE us_login = '" & VTlogin & "'"
    Set rstLogin = dbErp.OpenRecordset(strT)
    If Not rstLogin.EOF Then
        MsgBox "El usuario ya existe", vbCritical, "Login"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_user I, null,'" & txtDato(1).Text & "','" & txtDato(2).Text & "'"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VTCodigo = rstLogin(0)
    txtDato(0).Text = VTCodigo
    'Generacion del login
    Command_SQLx = "sp_addlogin " & txtDato(1).Text & "," & txtDato(2).Text & ", db_erp"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    'Creacion del usuario de la base
    Command_SQLx = "sp_adduser '" & txtDato(1).Text & "','" & txtDato(1).Text & "'"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    'Permisos a la base
    Command_SQLx = "sp_addrolemember 'db_owner', '" & txtDato(1).Text & "'"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Usuario creado correctamente", vbOKOnly, "Login"
    Screen.MousePointer = 0
    Exit Sub
ErrorInsertar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Insertar")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al insertar el usuario", vbCritical, "Insertar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLModificar()
Dim rstLogin As Recordset
Dim errObj As Error
Dim estado As String
Dim esta_anterior As String
Dim pass_anterior As String

If txtDato(0).Text = "" Then
    MsgBox "Debe ingresar el código del usuario", vbCritical, "Warning"
    txtDato(1).SetFocus
    Exit Sub
End If
If txtDato(1).Text = "" Then
    MsgBox "Debe ingresar el login del usuario", vbCritical, "Warning"
    txtDato(1).SetFocus
    Exit Sub
End If
On Error GoTo ErrorModificar
    Screen.MousePointer = 11
    strT = "SELECT us_password, us_estado FROM er_user WHERE us_codigo = " & txtDato(0).Text
    Set rstLogin = dbErp.OpenRecordset(strT)
    If Not rstLogin.EOF Then
        pass_anterior = rstLogin(0)
        esta_anterior = rstLogin(1)
    Else
        MsgBox "El usuario a eliminar no existe", vbCritical, "Login"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If chkEstado.Value = 1 Then
        estado = "A"
    Else
        estado = "I"
    End If
    If esta_anterior = estado Then
        Screen.MousePointer = 0
        MsgBox "No existe ninguna modificación"
        Exit Sub
    End If
    If txtDato(2).Text = "" Or pass_anterior = txtDato(2).Text Then
        Command_SQLx = "sp_user U, null," & txtDato(1).Text & ", null, " & estado & ", null, N"
        Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    Else
        Screen.MousePointer = 0
        MsgBox "Utilize la opción cambio de password"
        Exit Sub
    End If
    Screen.MousePointer = 0
    MsgBox "El usuario ha sido modificado correctamente", vbOKOnly, "Login"
    Exit Sub
ErrorModificar:
   VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Modificar")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al modificar el usuario", vbCritical, "Modificar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstLogin As Recordset
Dim errObj As Error

If txtDato(0).Text = "" Then
    MsgBox "Debe ingresar el código del usuario", vbCritical, "Warning"
    txtDato(1).SetFocus
    Exit Sub
End If

VT% = MsgBox("Esta seguro desea eliminar el usuario " & txtDato(0).Text & " ?", vbYesNo, "Login")
If VT% = 7 Then
    Exit Sub
End If

On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    strT = "SELECT us_codigo FROM er_user WHERE us_codigo = " & txtDato(0).Text
    Set rstLogin = dbErp.OpenRecordset(strT)
    If rstLogin.EOF Then
        MsgBox "El usuario a eliminar no existe", vbCritical, "Login"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_user D, null," & txtDato(1).Text & ", null, null, null, N"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
     Command_SQLx = "sp_dropuser " & txtDato(1).Text
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    Command_SQLx = "sp_droplogin " & txtDato(1).Text
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "El usuario ha sido eliminado correctamente", vbOKOnly, "Login"
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
        Var% = MsgBox("Error al eliminar el usuario", vbCritical, "Eliminar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLLimpiar()
    For i% = 0 To 3
        txtDato(i%).Text = ""
    Next i%
End Sub

Private Sub jmbBuscar_Click()
    Fbuscaruser.Show 1
End Sub

Private Sub jmbEliminar_Click()
    PLEliminar
End Sub

Private Sub jmbIngresar_Click()
     PLIngreso
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

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii% <> 8) Then
        If Len(txtDato(Index).Text) > 14 Then
            KeyAscii% = 0
        End If
    End If
End Sub
