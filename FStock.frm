VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FStock 
   Caption         =   "Stock Inicial"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   9795
   Begin VB.Frame frmbotones 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   8445
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
         TabIndex        =   1
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
         Picture         =   "FStock.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   5520
         TabIndex        =   2
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
         Picture         =   "FStock.frx":059A
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   6960
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
         Picture         =   "FStock.frx":0B34
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
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8415
      Begin VB.TextBox txtArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1905
         MaxLength       =   64
         TabIndex        =   5
         Top             =   300
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskCantidad 
         Height          =   300
         Left            =   1905
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblUnidad 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3255
         TabIndex        =   10
         Top             =   600
         Width           =   855
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
         Left            =   105
         TabIndex        =   8
         Top             =   300
         Width           =   1710
      End
      Begin VB.Label lblCantidad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
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
         TabIndex        =   7
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3255
         TabIndex        =   6
         Top             =   300
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VLPaso%                             'controla el cambio de los campos

Private Sub PLIngresar()
Dim rstStock As Recordset
Dim errObj As Error
Dim VTNow As String
Dim VTTime As String
'Validacion de datos
    If txtArticulo.Text = "" Then
        MsgBox "Debe ingresar el código del artículo", vbCritical, "Warning"
        Exit Sub
    End If
On Error GoTo ErrorInsertar
    Screen.MousePointer = 11
    strT = "SELECT st_articulo FROM er_stock_articulo WHERE st_articulo = " & Trim$(txtArticulo.Text)
    Set rstStock = dbErp.OpenRecordset(strT)
    If Not rstStock.EOF Then
        MsgBox "El Stock ya existe, no se puede ingresar", vbCritical, "Ingresar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If mskCantidad.Text = "" Then
        mskCantidad.Text = 0#
    End If
    VTNow$ = Format$(Now, VGFormatoFecha)
    VTTime$ = Format$(Time, "hh:mm:ss")
    Command_SQLx = "sp_stock @i_operacion = 'I', @i_articulo = " & txtArticulo.Text & _
    ", @i_fecha_tran = '" & VTNow$ & "',  @i_hora_tran = '" & VTTime$ & "', @i_cactual = " & mskCantidad.Text & _
    ", @i_formato_fecha = " & FMFormatoFecha(VGFormatoFecha) & _
    ", @i_login = '" & VGLOGIN & "', @i_transac = 'Stock Inicial'"
    Set rstStock = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Stock ingresado correctamente", vbOKOnly, "Ingreso"
    mskCantidad.Enabled = False
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
        MsgBox "Error al insertar el Stock", vbCritical, "Insertar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim rstUnidad As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    txtArticulo.Text = VGArticulo
    If txtArticulo.Text <> "" Then
        VGPaso% = False
        txtArticulo_LostFocus
         strT = "SELECT ar_unidad FROM er_articulo WHERE ar_codigo = " & Trim$(txtArticulo.Text)
        Set rstUnidad = dbErp.OpenRecordset(strT)
        If rstUnidad.EOF Then
            lblUnidad.Caption = ""
        Else
            lblUnidad.Caption = rstUnidad(0)
        End If
        strT = "SELECT st_cactual FROM er_stock_articulo WHERE st_articulo = " & Trim$(txtArticulo.Text)
        Set rstUnidad = dbErp.OpenRecordset(strT)
        If Not rstUnidad.EOF Then
            mskCantidad.Text = rstUnidad(0)
            mskCantidad.Enabled = False
            jmbIngresar.Enabled = False
        End If
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar el artículo", vbCritical, "Buscar")
    End If
End Sub

Private Sub jmbIngresar_Click()
    PLIngresar
End Sub

Private Sub jmbLimpiar_Click()
    txtArticulo.Text = ""
    lblDescripcion.Caption = ""
    lblUnidad.Caption = ""
    mskCantidad.Text = ""
    jmbIngresar.Enabled = True
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub txtarticulo_Change()
    VLPaso% = False
End Sub

Private Sub txtarticulo_GotFocus()
    VLPaso% = True
End Sub

Private Sub txtArticulo_LostFocus()
Dim rstArticulo As Recordset
Dim errObj As Error

    If VLPaso Then Exit Sub
    If txtArticulo.Text = "" Then Exit Sub
    On Error GoTo ErrorBuscar
    strT = "SELECT ar_descripcion FROM er_articulo WHERE ar_codigo = " & Trim$(txtArticulo.Text)
    Set rstArticulo = dbErp.OpenRecordset(strT)
    If rstArticulo.EOF Then
        lblDescripcion.Caption = ""
    Else
        lblDescripcion.Caption = rstArticulo(0)
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar Stock")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar la Stock", vbCritical, "Buscar Stock")
    End If
    Screen.MousePointer = 0
End Sub
