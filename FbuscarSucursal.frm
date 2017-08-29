VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FbuscarSucursal 
   Caption         =   "Buscar Sucursal"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid grdSucursal 
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
      Picture         =   "FbuscarSucursal.frx":0000
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
      Picture         =   "FbuscarSucursal.frx":08DA
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
      Picture         =   "FbuscarSucursal.frx":11B4
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
      Picture         =   "FbuscarSucursal.frx":1A8E
      Caption         =   "Escoger"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FbuscarSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Configura el grid empleados
    PLConfigura_grid
    jmbSiguientes.Enabled = False
    PLbuscar_sucursal
End Sub

Private Sub PLbuscar_sucursal()
Dim rstSucursal As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    Command_SQLx = "sp_sucursal @i_operacion ='S', @i_modo =0"
    Set rstSucursal = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    i = 0
    Do Until rstSucursal.EOF
        i = i + 1
        If i = 21 Then
            Exit Sub
        End If
        grdSucursal.Rows = grdSucursal.Rows + 1
        grdSucursal.Row = i
        grdSucursal.Col = 1
        grdSucursal.Text = rstSucursal(0)
        grdSucursal.Col = 2
        grdSucursal.Text = rstSucursal(1)
        grdSucursal.Col = 3
        grdSucursal.Text = rstSucursal(2)
        grdSucursal.Col = 4
        grdSucursal.Text = rstSucursal(3)
        grdSucursal.Col = 5
        grdSucursal.Text = rstSucursal(4)
        rstSucursal.MoveNext
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
        Var% = MsgBox("Error al buscar la sucursal", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLbuscar_siguientes()
Dim rstSucursal As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    If grdSucursal.Rows >= 20 Then
            grdSucursal.Row = grdSucursal.Rows - 1
            grdSucursal.Col = 1
            Command_SQLx = "sp_sucursals @i_operacion ='S', @i_modo = 1, @i_codigo = " & grdSucursal.Text
            Set rstSucursal = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
            i = 0
            Do Until rstSucursal.EOF
                 i = i + 1
                 If i = 21 Then
                     Exit Sub
                 End If
                 grdSucursal.Rows = grdSucursal.Rows + 1
                 grdSucursal.Row = i
                 grdSucursal.Col = 1
                 grdSucursal.Text = rstSucursal(0)
                 grdSucursal.Col = 2
                 grdSucursal.Text = rstSucursal(1)
                 grdSucursal.Col = 3
                 grdSucursal.Text = rstSucursal(2)
                 grdSucursal.Col = 4
                 grdSucursal.Text = rstSucursal(3)
                 grdSucursal.Col = 5
                 grdSucursal.Text = rstSucursal(4)
                 rstSucursal.MoveNext
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
        Var% = MsgBox("Error al buscar la sucursal", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLConfigura_grid()
    grdSucursal.Cols = 6
    grdSucursal.Rows = 1
    grdSucursal.ColWidth(0) = 250
    grdSucursal.ColWidth(1) = 1000
    grdSucursal.ColWidth(2) = 1500
    grdSucursal.ColWidth(3) = 1000
    grdSucursal.ColWidth(4) = 500
    grdSucursal.ColWidth(5) = 500
    grdSucursal.Row = 0
    grdSucursal.Col = 1
    grdSucursal.Text = "Código"
    grdSucursal.Col = 2
    grdSucursal.Text = "Descripción"
    grdSucursal.Col = 3
    grdSucursal.Text = "Dirección"
    grdSucursal.Col = 4
    grdSucursal.Text = "Pto Emisión"
    grdSucursal.Col = 5
    grdSucursal.Text = "Pto Facturación"
End Sub

Private Sub PLEscoger()
Dim rstSucursal As Recordset
Dim errObj As Error
Dim VTTidentidad As String
Dim VTCedula As String
Dim VTFecha_ing As String
Dim VTEstado As String

On Error GoTo ErrorEscoger
    VTRow% = grdSucursal.Row
    If VTRow% = 0 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    grdSucursal.Col = 1
    FSucursal.txtCodigo.Text = grdSucursal.Text
    grdSucursal.Col = 2
    FSucursal.txtDescripcion.Text = grdSucursal.Text
    grdSucursal.Col = 3
    FSucursal.txtDireccion.Text = grdSucursal.Text
    grdSucursal.Col = 4
    FSucursal.txtPtoemision.Text = grdSucursal.Text
    grdSucursal.Col = 5
    FSucursal.txtPtofacturacion.Text = grdSucursal.Text
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
        Var% = MsgBox("Error al buscar la sucursal", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub grdSucursal_Click()
    PMLineaG grdSucursal
End Sub

Private Sub grdSucursal_DblClick()
    PLEscoger
End Sub

Private Sub jmbbuscar_Click()
    PLbuscar_sucursal
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
