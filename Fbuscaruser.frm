VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form Fbuscaruser 
   Caption         =   "Buscar Usuarios"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid grdusuarios 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   4260
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin JMButton.JMBcontrol jmbSiguientes 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
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
      Picture         =   "Fbuscaruser.frx":0000
      Caption         =   "Siguientes"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbbuscar 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
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
      Picture         =   "Fbuscaruser.frx":08DA
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbSalir 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2760
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
      Picture         =   "Fbuscaruser.frx":11B4
      Caption         =   "Salir"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbEscoger 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
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
      Picture         =   "Fbuscaruser.frx":1A8E
      Caption         =   "Escoger"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "Fbuscaruser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Configura el grid usuarios
    PLConfigura_grid
    jmbSiguientes.Enabled = False
    PLbuscar_usuarios
End Sub

Private Sub PLbuscar_usuarios()
Dim rstUser As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    Command_SQLx = "sp_user @i_operacion = 'S', @i_modo = 0"
    Set rstUser = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    i = 0
    Do Until rstUser.EOF
        i = i + 1
        If i = 21 Then
            Exit Sub
        End If
        grdusuarios.Rows = grdusuarios.Rows + 1
        grdusuarios.Row = i
        grdusuarios.Col = 1
        grdusuarios.Text = rstUser(0)
        grdusuarios.Col = 2
        grdusuarios.Text = rstUser(1)
        grdusuarios.Col = 3
        grdusuarios.Text = rstUser(2)
        rstUser.MoveNext
    Loop
    If i = 21 Then
        jmbSiguientes.Enabled = True
        jmbbuscar.Enabled = False
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
        Var% = MsgBox("Error al buscar el registro", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLBuscar_Siguientes()
Dim rstUser As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
     If grdusuarios.Rows >= 20 Then
            grdusuarios.Row = grdusuarios.Rows - 1
            grdusuarios.Col = 1
            Command_SQLx = "sp_user @i_operacion = 'S', @i_modo = 1, @i_codigo = " & grdusuarios.Text
            Set rstUser = dbNomina.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
            i = 0
            Do Until rstUser.EOF
                 i = i + 1
                 If i = 21 Then
                     Exit Sub
                 End If
                 grdusuarios.Rows = grdusuarios.Rows + 1
                 grdusuarios.Row = i
                 grdusuarios.Col = 1
                 grdusuarios.Text = rstUser(0)
                 grdusuarios.Col = 2
                 grdusuarios.Text = rstUser(1)
                 grdusuarios.Col = 3
                 grdusuarios.Text = rstUser(2)
                 rstUser.MoveNext
            Loop
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
        Var% = MsgBox("Error al buscar el registro", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLConfigura_grid()
    grdusuarios.Cols = 4
    grdusuarios.Rows = 1
    grdusuarios.ColWidth(0) = 250
    grdusuarios.ColWidth(1) = 800
    grdusuarios.ColWidth(2) = 10000
    grdusuarios.Row = 0
    grdusuarios.Col = 1
    grdusuarios.Text = "Código"
    grdusuarios.Col = 2
    grdusuarios.Text = "Login"
    grdusuarios.Col = 3
    grdusuarios.Text = "Estado"
End Sub

Private Sub PLEscoger()
    VTRow% = grdusuarios.Row
    If VTRow% = 0 Then
        Exit Sub
    End If
    grdusuarios.Col = 1
    Fusuario.txtDato(0).Text = grdusuarios.Text
    grdusuarios.Col = 2
    Fusuario.txtDato(1).Text = grdusuarios.Text
    grdusuarios.Col = 3
    If grdusuarios.Text = "A" Then
            Fusuario.chkEstado.Value = 1
    Else
            Fusuario.chkEstado.Value = 0
    End If
    Unload Me
End Sub

Private Sub grdusuarios_Click()
    PMLineaG grdusuarios
End Sub

Private Sub grdusuarios_DblClick()
    PLEscoger
End Sub

Private Sub jmbBuscar_Click()
    PLbuscar_usuarios
End Sub

Private Sub jmbEscoger_Click()
    PLEscoger
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub jmbSiguientes_Click()
    PLBuscar_Siguientes
End Sub
