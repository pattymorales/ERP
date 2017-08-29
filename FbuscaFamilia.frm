VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FbuscaFamilia 
   Caption         =   "Buscar Familia"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid grdFamilia 
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
      Picture         =   "FbuscaFamilia.frx":0000
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
      Picture         =   "FbuscaFamilia.frx":08DA
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
      Picture         =   "FbuscaFamilia.frx":11B4
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
      Picture         =   "FbuscaFamilia.frx":1A8E
      Caption         =   "Escoger"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FbuscaFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Configura el grid empleados
    PLConfigura_grid
    jmbSiguientes.Enabled = False
    PLbuscar_familia
End Sub

Private Sub PLbuscar_familia()
Dim rstFamilia As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    Command_SQLx = "sp_familia @i_operacion ='S', @i_modo =0"
    Set rstFamilia = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    i = 0
    Do Until rstFamilia.EOF
        i = i + 1
        If i = 21 Then
            Exit Sub
        End If
        grdFamilia.Rows = grdFamilia.Rows + 1
        grdFamilia.Row = i
        grdFamilia.Col = 1
        grdFamilia.Text = rstFamilia(0)
        grdFamilia.Col = 2
        grdFamilia.Text = rstFamilia(1)
        grdFamilia.Col = 3
        grdFamilia.Text = rstFamilia(2)
        rstFamilia.MoveNext
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
        Var% = MsgBox("Error al buscar el empleado", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLbuscar_siguientes()
Dim rstFamilia As Recordset
Dim errObj As Error

On Error GoTo ErrorBuscar
    Screen.MousePointer = 11
    If grdFamilia.Rows >= 20 Then
            grdFamilia.Row = grdFamilia.Rows - 1
            grdFamilia.Col = 1
            Command_SQLx = "sp_familia @i_operacion ='S', @i_modo = 1, @i_codigo = " & grdFamilia.Text
            Set rstFamilia = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
            i = 0
            Do Until rstFamilia.EOF
                 i = i + 1
                 If i = 21 Then
                     Exit Sub
                 End If
                 grdFamilia.Rows = grdFamilia.Rows + 1
                 grdFamilia.Row = i
                 grdFamilia.Col = 1
                 grdFamilia.Text = rstFamilia(0)
                 grdFamilia.Col = 2
                 grdFamilia.Text = rstFamilia(1)
                 grdFamilia.Col = 3
                 grdFamilia.Text = rstFamilia(2)
                 grdFamilia.Col = 4
                 grdFamilia.Text = rstFamilia(3)
                 rstFamilia.MoveNext
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
        Var% = MsgBox("Error al buscar el empleado", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
End Sub

Private Sub PLConfigura_grid()
        grdFamilia.Cols = 4
        grdFamilia.Rows = 1
        grdFamilia.ColWidth(0) = 250
        grdFamilia.ColWidth(1) = 2000
        grdFamilia.ColWidth(2) = 1000
        grdFamilia.ColWidth(3) = 1000
        grdFamilia.Row = 0
        grdFamilia.Col = 1
        grdFamilia.Text = "Código"
        grdFamilia.Col = 2
        grdFamilia.Text = "Descripción"
        grdFamilia.Col = 3
        grdFamilia.Text = "Cuenta Contable"
End Sub

Private Sub PLEscoger()
Dim rstFamilia As Recordset
Dim errObj As Error
Dim VTTidentidad As String
Dim VTCedula As String
Dim VTFecha_ing As String
Dim VTEstado As String

On Error GoTo ErrorEscoger
    VTRow% = grdFamilia.Row
    If VTRow% = 0 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    If VGForma = "FFamilia" Then
        grdFamilia.Col = 1
        FFamilia.txtCodigo.Text = grdFamilia.Text
        grdFamilia.Col = 2
        FFamilia.txtDescripcion.Text = grdFamilia.Text
        grdFamilia.Col = 3
        FFamilia.txtCta_Contable.Text = grdFamilia.Text
    End If
    If VGForma = "FArticulo" Then
        grdFamilia.Col = 1
        FArticulo.txtFamilia.Text = grdFamilia.Text
        grdFamilia.Col = 2
        FArticulo.lbldesfamilia.Caption = grdFamilia.Text
    End If
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
        Var% = MsgBox("Error al escoger el empleado", vbCritical, "Buscar")
    End If
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub grdFamilia_Click()
    PMLineaG grdFamilia
End Sub

Private Sub grdFamilia_DblClick()
    PLEscoger
End Sub

Private Sub jmbbuscar_Click()
    PLbuscar_familia
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
