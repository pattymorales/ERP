VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FbuscarArticulo 
   Caption         =   "Buscar Artículo"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid grdArticulos 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
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
      Picture         =   "FbuscaArticulo.frx":0000
      Caption         =   "Siguientes"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbbuscar 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
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
      Picture         =   "FbuscaArticulo.frx":08DA
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbSalir 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   2640
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
      Picture         =   "FbuscaArticulo.frx":11B4
      Caption         =   "Salir"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbEscoger 
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
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
      Picture         =   "FbuscaArticulo.frx":1A8E
      Caption         =   "Escoger"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FbuscarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Configura el grid empleados
    PLConfigura_grid
    jmbSiguientes.Enabled = False
    PLbuscar_articulos
End Sub

Private Sub PLbuscar_articulos()
Dim rstarticulos As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    Command_SQLx = "sp_articulo @i_operacion ='S', @i_modo =0"
    Set rstarticulos = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    i = 0
    Do Until rstarticulos.EOF
        i = i + 1
        If i = 21 Then
            Exit Sub
        End If
        grdArticulos.Rows = grdArticulos.Rows + 1
        grdArticulos.Row = i
        grdArticulos.Col = 1
        grdArticulos.Text = rstarticulos(0)
        grdArticulos.Col = 2
        grdArticulos.Text = rstarticulos(1)
        grdArticulos.Col = 3
        grdArticulos.Text = rstarticulos(2)
        grdArticulos.Col = 4
        grdArticulos.Text = rstarticulos(3)
        grdArticulos.Col = 5
        grdArticulos.Text = rstarticulos(4)
        rstarticulos.MoveNext
    Loop
    If i = 21 Then
        jmbSiguientes.Enabled = True
        jmbBuscar.Enabled = False
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
        Var% = MsgBox("Error al buscar el artículo", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLbuscar_siguientes()
Dim rstarticulos As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    If grdArticulos.Rows >= 20 Then
            grdArticulos.Row = grdArticulos.Rows - 1
            grdArticulos.Col = 1
            Command_SQLx = "sp_articulos @i_operacion ='S', @i_modo = 1, @i_codigo = " & grdArticulos.Text
            Set rstarticulos = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
            i = 0
            Do Until rstarticulos.EOF
                 i = i + 1
                 If i = 21 Then
                     Exit Sub
                 End If
                 grdArticulos.Rows = grdArticulos.Rows + 1
                 grdArticulos.Row = i
                 grdArticulos.Col = 1
                 grdArticulos.Text = rstarticulos(0)
                 grdArticulos.Col = 2
                 grdArticulos.Text = rstarticulos(1)
                 grdArticulos.Col = 3
                 grdArticulos.Text = rstarticulos(2)
                 grdArticulos.Col = 4
                 grdArticulos.Text = rstarticulos(3)
                 grdArticulos.Col = 5
                 grdArticulos.Text = rstarticulos(4)
                 rstarticulos.MoveNext
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
        Var% = MsgBox("Error al buscar el artículo", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLConfigura_grid()
        grdArticulos.Cols = 6
        grdArticulos.Rows = 1
        grdArticulos.ColWidth(0) = 250
        grdArticulos.ColWidth(1) = 1000
        grdArticulos.ColWidth(2) = 1500
        grdArticulos.ColWidth(3) = 1000
        grdArticulos.ColWidth(4) = 1000
        grdArticulos.ColWidth(5) = 500
        grdArticulos.Row = 0
        grdArticulos.Col = 1
        grdArticulos.Text = "Código"
        grdArticulos.Col = 2
        grdArticulos.Text = "Descripción"
        grdArticulos.Col = 3
        grdArticulos.Text = "Familia"
        grdArticulos.Col = 4
        grdArticulos.Text = "Costo"
        grdArticulos.Col = 5
        grdArticulos.Text = "Unidad"
End Sub

Private Sub PLEscoger()
Dim rstarticulos As Recordset
Dim errObj As Error
Dim VTTidentidad As String
Dim VTCedula As String
Dim VTFecha_ing As String
Dim VTEstado As String

On Error GoTo ErrorEscoger
    VTRow% = grdArticulos.Row
    If VTRow% = 0 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    grdArticulos.Col = 1
    FArticulo.txtCodigo.Text = grdArticulos.Text
    grdArticulos.Col = 2
    FArticulo.txtDescripcion.Text = grdArticulos.Text
    grdArticulos.Col = 3
    FArticulo.txtFamilia.Text = grdArticulos.Text
    grdArticulos.Col = 4
    FArticulo.mskCosto.Text = grdArticulos.Text
    grdArticulos.Col = 5
    FArticulo.txtUnidad.Text = grdArticulos.Text
    Screen.MousePointer = 0
    Unload Me
    Exit Sub
ErrorEscoger:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Buscar")
         VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al buscar el artículo", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub grdArticulos_Click()
    PMLineaG grdArticulos
End Sub

Private Sub grdArticulos_DblClick()
    PLEscoger
End Sub

Private Sub jmbbuscar_Click()
    PLbuscar_articulos
End Sub

Private Sub jmbEscoger_Click()
    PLEscoger
End Sub

Private Sub jmbSalir_Click()
    Unload Me
End Sub

Private Sub jmbSiguientes_Click()
    PLbuscar_siguientes
End Sub
