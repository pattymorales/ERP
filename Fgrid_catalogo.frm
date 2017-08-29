VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form Fgrid_catalogo 
   Caption         =   "Catálogo"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid grdcatalogo 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   5741
      _StockProps     =   77
      BackColor       =   16777215
   End
   Begin JMButton.JMBcontrol jmbsiguientes 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
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
      Picture         =   "Fgrid_catalogo.frx":0000
      Caption         =   "Siguientes"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbbuscar 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3480
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
      Picture         =   "Fgrid_catalogo.frx":08DA
      Caption         =   "Buscar"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbSalir 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3480
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
      Picture         =   "Fgrid_catalogo.frx":11B4
      Caption         =   "Salir"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbEscoger 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3480
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
      Picture         =   "Fgrid_catalogo.frx":1A8E
      Caption         =   "Escoger"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "Fgrid_catalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Configura el grid
    PLConfigura_grid
    'jmbboton(1).Enabled = False
    PLbuscar_datos
End Sub

Private Sub PLConfigura_grid()
    grdcatalogo.Cols = 3
    grdcatalogo.Rows = 21
    grdcatalogo.ColWidth(0) = 250
    grdcatalogo.ColWidth(1) = 800
    grdcatalogo.ColWidth(2) = 2600
    grdcatalogo.Row = 0
    grdcatalogo.Col = 1
    grdcatalogo.Text = "Código"
    grdcatalogo.Col = 2
    grdcatalogo.Text = "Descripción"
End Sub

Private Sub PLbuscar_datos()
Dim rstCatalogo As Recordset
On Error GoTo ErrorBuscar
    grdcatalogo.Rows = 1
    grdcatalogo.Rows = 21
    Screen.MousePointer = 11
    Command_SQLx = StoredProcedure & " @i_operacion = 'S', @i_modo = 0"
    Set rstCatalogo = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    i = 0
    Do Until rstCatalogo.EOF
        i = i + 1
        If i = 21 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        grdcatalogo.Row = i
        If rstCatalogo.Fields(0).Value <> "" Then
            grdcatalogo.Col = 1
            grdcatalogo.Text = rstCatalogo(0)
        End If
        If rstCatalogo.Fields(1).Value <> "" Then
            grdcatalogo.Col = 2
            grdcatalogo.Text = rstCatalogo(1)
        End If
        rstCatalogo.MoveNext
    Loop
    If i = 21 Then
        jmbSiguientes.Enabled = True
        jmbBuscar.Enabled = False
    End If
    grdcatalogo.Rows = i + 1
    Screen.MousePointer = 0
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        MsgBox errObj.Number & " " & errObj.Description, vbCritical, "Buscar"
         VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al buscar el usuario", vbCritical, "Buscar"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub jmbbuscar_Click()
    PLbuscar_datos
End Sub

Private Sub jmbEscoger_Click()
     grdcatalogo_DblClick
End Sub

Private Sub jmbSalir_Click()
    ReDim Temporales(10)
    If grdcatalogo.Rows > 1 Then
        PMLimpiaG grdcatalogo
    End If
    Unload Me
End Sub

Sub grdcatalogo_Click()
    PMLineaG grdcatalogo
End Sub

Sub grdcatalogo_DblClick()
        If grdcatalogo.Row = 0 Then
                Exit Sub
        Else
            VTRow% = grdcatalogo.Row
            If grdcatalogo.Rows > 1 Then
                 grdcatalogo.Row = 1
                 grdcatalogo.Col = 1
                 'verifica que el grid no esté vacío
                 If grdcatalogo.Text <> "" Then
                     grdcatalogo.Row = VTRow%
                     ReDim Temporales(grdcatalogo.Cols - 1)
                     For c% = 1 To grdcatalogo.Cols - 1
                          grdcatalogo.Col = c%
                          Temporales(c%) = grdcatalogo.Text
                     Next c%
                 Else
                     ReDim Temporales(10)
                 End If
                 PMLimpiaG grdcatalogo
            Else
                 ReDim Temporales(10)
            End If
            FPrincipal.pnlHelpLine.Caption = ""
            FPrincipal.pnlTransaccionLine.Caption = ""
            Unload Me
        End If
End Sub

Sub grdcatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
    PMLineaG grdcatalogo
End Sub

Sub grdcatalogo_KeyPress(KeyAscii As Integer)
'*********************************************************
'Objetivo:  Realiza la función de los botones escoger y
'           cancelar cuando sobre el grid se digita {ENTER}
'           o {ESC}
'Input   :  KeyAscii        código de la tecla presionada
'Output  :  ninguno
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Nov/93  M.Davila            Emisión Inicial
'*********************************************************
    If KeyAscii = 13 Then
        'chequeo que el grid no esté vacío
        VTRow% = grdcatalogo.Row
        grdcatalogo.Row = 1
        grdcatalogo.Col = 1
        If grdcatalogo.Text <> "" Then
            grdcatalogo.Row = VTRow%
            ReDim Temporales(grdcatalogo.Cols)
            For c% = 1 To grdcatalogo.Cols - 1
                grdcatalogo.Col = c%
                Temporales(c%) = grdcatalogo.Text
            Next c%
        Else
            ReDim Temporales(10)
        End If
        PMLimpiaG grdcatalogo
        grdcatalogo.Hide
    Else
        If KeyAscii = KEY_ESCAPE Then
            ReDim Temporales(10)
            PMLimpiaG grdcatalogo
            grdcatalogo.Hide
        End If
    End If
End Sub

Sub grdcatalogo_KeyUp(KeyCode As Integer, Shift As Integer)
    PMLineaG grdcatalogo
End Sub

Private Sub jmbSiguientes_Click()
Dim rstCatalogo As Recordset

On Error GoTo ErrorBuscar
    If grdcatalogo.Rows >= CGMaximoRows% Then
        If VGPSiguiente% > 0 Then
            grdcatalogo.Col = 1
            grdcatalogo.Row = grdcatalogo.Rows - 1
            Screen.MousePointer = 11
            Command_SQLx = StoredProcedure & " S,'" & grdcatalogo.Text & "', null , 1"
            Set rstCatalogo = dbNomina.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
            i = 0
            Do Until rstCatalogo.EOF
                i = i + 1
                If i = 21 Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                grdcatalogo.Row = i
                If rstCatalogo.Fields(0).Value <> "" Then
                    grdcatalogo.Col = 1
                    grdcatalogo.Text = rstCatalogo(0)
                End If
                If rstCatalogo.Fields(1).Value <> "" Then
                    grdcatalogo.Col = 2
                    grdcatalogo.Text = rstCatalogo(1)
                End If
                rstCatalogo.MoveNext
            Loop
            If i = 21 Then
                jmbSiguientes.Enabled = True
                jmbBuscar.Enabled = False
            End If
            grdcatalogo.Rows = i + 1
            Screen.MousePointer = 0
            If grdcatalogo.Tag >= CGMaximoRows% - 1 Then
                jmbSiguientes.Enabled = True
            Else
                jmbSiguientes.Enabled = False
            End If
            grdcatalogo.TopRow = grdcatalogo.Rows - CGMaximoRows%
        End If
    End If
    Exit Sub
ErrorBuscar:
    VGdisperror = False
    For Each errObj In Errors
        MsgBox errObj.Number & " " & errObj.Description, vbCritical, "Buscar"
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al buscar el usuario", vbCritical, "Buscar"
    End If
    Screen.MousePointer = 0
End Sub
