VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FArticulo 
   Caption         =   "Artículo"
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
      Left            =   30
      TabIndex        =   17
      Top             =   2160
      Width           =   7935
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
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
         Picture         =   "Farticulo.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbModificar 
         Height          =   495
         Left            =   1560
         TabIndex        =   7
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
         Picture         =   "Farticulo.frx":059A
         Caption         =   "Modificar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbEliminar 
         Height          =   495
         Left            =   3000
         TabIndex        =   8
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
         Picture         =   "Farticulo.frx":0E74
         Caption         =   "Eliminar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   4440
         TabIndex        =   9
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
         Picture         =   "Farticulo.frx":174E
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   5880
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
         Picture         =   "Farticulo.frx":1CE8
         Caption         =   "Salir"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos"
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
      Height          =   2055
      Left            =   20
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtUnidad 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   20
         Top             =   1560
         Width           =   1320
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   64
         TabIndex        =   1
         Top             =   660
         Width           =   5955
      End
      Begin VB.TextBox txtFamilia 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   1320
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   64
         TabIndex        =   0
         Top             =   360
         Width           =   1320
      End
      Begin MSMask.MaskEdBox mskCosto 
         Height          =   300
         Left            =   1920
         TabIndex        =   5
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdBfamilia 
         Height          =   285
         Left            =   3260
         TabIndex        =   3
         Tag             =   "21315"
         ToolTipText     =   "Buscar"
         Top             =   960
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Farticulo.frx":25C2
      End
      Begin Threed.SSCommand cmdBunidad 
         Height          =   285
         Left            =   3260
         TabIndex        =   21
         Tag             =   "21315"
         ToolTipText     =   "Buscar"
         Top             =   1560
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Farticulo.frx":2E9C
      End
      Begin VB.Label lblDesunidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3690
         TabIndex        =   22
         Top             =   1560
         Width           =   4170
      End
      Begin VB.Label lblunidad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
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
         TabIndex        =   19
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label lblcodigo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1710
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
         TabIndex        =   15
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblFamilia 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Familia de Artículo:"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lblCosto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Costo del Artículo:"
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1605
      End
      Begin VB.Label lbldesfamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3690
         TabIndex        =   4
         Top             =   960
         Width           =   4170
      End
   End
   Begin JMButton.JMBcontrol jmbStock 
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   840
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
      Picture         =   "Farticulo.frx":3776
      Caption         =   "Stock"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbbuscar 
      Height          =   495
      Left            =   8040
      TabIndex        =   18
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
      Picture         =   "Farticulo.frx":4050
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VLPaso%                             'controla el cambio de los campos

Private Sub PLIngresar()
Dim rstArticulo As Recordset
Dim errObj As Error
Dim VTCodigo As String

'Validacion de datos
    If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción del artículo", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If txtFamilia.Text = "" Then
        MsgBox "Debe ingresar la familia del artículo", vbCritical, "Warning"
        txtFamilia.SetFocus
        Exit Sub
    End If
    If txtUnidad.Text = "" Then
        MsgBox "Debe ingresar la unidad del artículo", vbCritical, "Warning"
        txtUnidad.SetFocus
        Exit Sub
    End If

On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    Command_SQLx = "sp_articulo @i_operacion = 'I', @i_descripcion = " & _
    "'" & txtDescripcion.Text & "', @i_familia = '" & txtFamilia.Text & "'," & _
    " @i_costo  = " & mskCosto.Text & ",  @i_unidad = '" & txtUnidad.Text & "'"
    Set rstArticulo = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VTCodigo = rstArticulo(0)
    txtCodigo.Text = VTCodigo
    MsgBox "Artículo creado correctamente", vbOKOnly, "Ingreso"
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
        MsgBox "Error al insertar el artículo", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbboton_Click(Index As Integer)
    FStock.Show 1
End Sub

Private Sub cmdBfamilia_Click()
    VGForma = "FArticulo"
    FbuscaFamilia.Show 1
    If txtFamilia.Text <> "" Then
        VLPaso% = False
        txtFamilia_LostFocus
    Else
        VLPaso% = False
    End If
End Sub

Private Sub cmdBunidad_Click()
    StoredProcedure = "sp_unidades"
    Fgrid_catalogo.Show 1
    If Temporales(1) <> "" Then
        txtUnidad.Text = Temporales(1)
        lblDesunidad.Caption = Temporales(2)
        VLPaso% = False
        txtUnidad_LostFocus
    Else
        VLPaso% = False
    End If
End Sub

Private Sub jmbbuscar_Click()
    FbuscarArticulo.Show 1
    If txtFamilia.Text <> "" Then
        VLPaso% = False
        txtFamilia_LostFocus
        txtUnidad_LostFocus
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
Dim rstArticulo As Recordset
Dim errObj As Error

'Validacion de datos
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código del artículo", vbCritical, "Warning"
        Exit Sub
    End If
     If txtDescripcion.Text = "" Then
        MsgBox "Debe ingresar la descripción del artículo", vbCritical, "Warning"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If txtFamilia.Text = "" Then
        MsgBox "Debe ingresar la familia del artículo", vbCritical, "Warning"
        txtFamilia.SetFocus
        Exit Sub
    End If
On Error GoTo ErrorModificar
Screen.MousePointer = 11
        Command_SQLx = "sp_articulo @i_operacion = 'U', @i_codigo = '" & _
        txtCodigo.Text & "', @i_descripcion = '" & txtDescripcion.Text & "'," & _
        " @i_familia  = '" & txtFamilia.Text & "', @i_costo = " & mskCosto.Text & _
        ", @i_unidad = '" & txtUnidad.Text & "'"
        Set rstArticulo = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
        MsgBox "Artículo modificado correctamente", vbOKOnly, "Modificar"
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
        Var% = MsgBox("Error al modificar el artículo", vbCritical, "Modificar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLEliminar()
Dim rstArticulo As Recordset
Dim errObj As Error
        
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el código del artículo", vbCritical, "Warning"
        Exit Sub
    End If
    VT% = MsgBox("Esta seguro de eliminar el artículo:  " & txtCodigo.Text & " ?", vbYesNo, "Eliminar")
    If VT% = 7 Then
        Exit Sub
    End If
On Error GoTo ErrorEliminar
    Screen.MousePointer = 11
    strT = "SELECT ar_codigo FROM er_articulo WHERE ar_codigo = " & Trim$(txtCodigo.Text)
    Set rstArticulo = dbErp.OpenRecordset(strT)
    If rstArticulo.EOF Then
        MsgBox "El artículo no existe", vbCritical, "Eliminar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command_SQLx = "sp_articulo @i_operacion = 'D', @i_codigo = '" & txtCodigo.Text & "'"
    Set rstArticulo = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "El artículo ha sido eliminado correctamente", vbOKOnly, "Eliminar"
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
        Var% = MsgBox("Error al eliminar el artículo", vbCritical, "Eliminar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLLimpiar()
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    txtFamilia.Text = ""
    lbldesfamilia.Caption = ""
    mskCosto.Text = ""
    txtUnidad.Text = ""
    lblDesunidad.Caption = ""
    jmbIngresar.Enabled = True
End Sub

Private Sub jmbStock_Click()
    If txtCodigo.Text = "" Then
        MsgBox "Debe escoger un artículo para ingresar el stock"
        Exit Sub
    End If
    VGArticulo = txtCodigo.Text
    FStock.Show 1
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtFamilia_Change()
    VLPaso% = False
End Sub

Private Sub txtFamilia_GotFocus()
    VLPaso% = True
End Sub

Private Sub txtFamilia_LostFocus()
Dim rstFamilia As Recordset
Dim errObj As Error

    If VLPaso Then Exit Sub
    If txtFamilia.Text = "" Then Exit Sub
    On Error GoTo ErrorBuscar
    strT = "SELECT fa_descripcion FROM er_familia WHERE fa_codigo = '" & Trim$(txtFamilia.Text) & "'"
    Set rstFamilia = dbErp.OpenRecordset(strT)
    If rstFamilia.EOF Then
        txtFamilia.Text = ""
        lbldesfamilia.Caption = ""
    Else
        lbldesfamilia.Caption = rstFamilia(0)
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar Familia")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar la familia", vbCritical, "Buscar Familia")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub txtUnidad_Change()
    VLPaso% = False
End Sub

Private Sub txtUnidad_GotFocus()
    VLPaso% = True
End Sub

Private Sub txtUnidad_LostFocus()
Dim rstUnidad As Recordset
Dim errObj As Error

    If VLPaso Then Exit Sub
    If txtUnidad.Text = "" Then Exit Sub
    On Error GoTo ErrorBuscar
    strT = "SELECT un_descripcion FROM er_unidades WHERE un_codigo = '" & Trim$(txtUnidad.Text) & "'"
    Set rstUnidad = dbErp.OpenRecordset(strT)
    If rstUnidad.EOF Then
        txtUnidad.Text = ""
        lblDesunidad.Caption = ""
    Else
        lblDesunidad.Caption = rstUnidad(0)
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar Unidad")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar la familia", vbCritical, "Buscar Unidad")
    End If
    Screen.MousePointer = 0
End Sub

