VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FCStock 
   Caption         =   "Corrección de Stock"
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
      TabIndex        =   9
      Top             =   1560
      Width           =   8685
      Begin JMButton.JMBcontrol jmbIngresar 
         Height          =   495
         Left            =   120
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
         Picture         =   "FCstock.frx":0000
         Caption         =   "Ingresar"
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbLimpiar 
         Height          =   495
         Left            =   5760
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
         Picture         =   "FCstock.frx":059A
         Caption         =   "Limpiar"
         CaptionPlace    =   4
         WordWrap        =   -1  'True
         Border          =   3
      End
      Begin JMButton.JMBcontrol jmbSalir 
         Height          =   495
         Left            =   7200
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
         Picture         =   "FCstock.frx":0B34
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   64
         TabIndex        =   1
         Top             =   300
         Width           =   1395
      End
      Begin MSMask.MaskEdBox mskCactual 
         Height          =   300
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCnueva 
         Height          =   300
         Left            =   1920
         TabIndex        =   3
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmdBfamilia 
         Height          =   285
         Left            =   3330
         TabIndex        =   4
         Tag             =   "21315"
         ToolTipText     =   "Buscar"
         Top             =   300
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
         Picture         =   "FCstock.frx":140E
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
      Begin VB.Label lblCantidadStock 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cantidad actual:"
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
         Width           =   1410
      End
      Begin VB.Label lblCnueva 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cantidad nueva:"
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
         TabIndex        =   6
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3770
         TabIndex        =   5
         Top             =   300
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FCStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VLPaso%                             'controla el cambio de los campos

Private Sub PLLimpiar()
    txtArticulo.Text = ""
    lblDescripcion.Caption = ""
    mskCactual.Text = 0
    mskCnueva.Text = 0
    mskCnueva.Enabled = True
    jmbIngresar.Enabled = True
End Sub
Private Sub cmdBfamilia_Click()
    StoredProcedure = "sp_articulo"
    Fgrid_catalogo.Show 1
    If Temporales(1) <> "" Then
        txtArticulo.Text = Temporales(1)
        lblDescripcion.Caption = Temporales(2)
        VLPaso% = False
        txtArticulo_LostFocus
        mskCnueva.Text = 0#
    Else
        VLPaso% = False
    End If
End Sub


Private Sub jmbIngresar_Click()
    PLIngresar
End Sub

Private Sub jmbLimpiar_Click()
    PLLimpiar
End Sub

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
    strT = "SELECT ar_codigo FROM er_articulo WHERE ar_codigo = " & Trim$(txtArticulo.Text)
    Set rstStock = dbErp.OpenRecordset(strT)
    If rstStock.EOF Then
        MsgBox "Artículo no existe, no se puede realizar la corrección de stock", vbCritical, "Ingresar"
        Screen.MousePointer = 0
        Exit Sub
    End If
    VTNow$ = Format$(Now, VGFormatoFecha)
    VTTime$ = Format$(Time, "hh:mm:ss")
    Command_SQLx = "sp_stock @i_operacion = 'U', @i_articulo = " & txtArticulo.Text & _
    ", @i_fecha_tran = '" & VTNow$ & "', @i_hora_tran = '" & VTTime$ & "', @i_cactual = " & mskCnueva.Text & _
    ", @i_formato_fecha = " & FMFormatoFecha(VGFormatoFecha) & _
    ", @i_login = '" & VGLOGIN & "', @i_transac = 'Correccion de Stock'"
    Set rstStock = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "Correción de Stock realizada correctamente", vbOKOnly, "Ingreso"
    mskCnueva.Enabled = False
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
Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub txtarticulo_Change()
    VLPaso% = False
End Sub

Private Sub txtarticulo_GotFocus()
    VLPaso% = True
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    If (KeyAscii% <> 8) And (KeyAscii% <> 32) And (KeyAscii% <> 39) And (KeyAscii% < 48 Or KeyAscii% > 57) Then
        KeyAscii% = 0
    Else
        KeyAscii% = Asc(UCase$(Chr$(KeyAscii%)))
    End If
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
        txtArticulo.Text = ""
        lblDescripcion.Caption = ""
    Else
        lblDescripcion.Caption = rstArticulo(0)
    End If
    If txtArticulo.Text <> "" Then
        PLBuscar_stock
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar artículo")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar la familia", vbCritical, "Buscar artículo")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLBuscar_stock()
Dim rstStock As Recordset
Dim errObj As Error

On Error GoTo ErrorbuscarStock
    strT = "SELECT st_cactual FROM er_stock_articulo WHERE st_articulo = " & Trim$(txtArticulo.Text)
    Set rstStock = dbErp.OpenRecordset(strT)
    If rstStock.EOF Then
        mskCactual.Text = 0
    Else
        mskCactual.Text = rstStock(0)
    End If
    Exit Sub
ErrorbuscarStock:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar Stock")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al insertar el Stock", vbCritical, "Buscar Stock"
    End If
    Screen.MousePointer = 0
End Sub
