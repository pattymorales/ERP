VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Begin VB.Form grid_valores 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Registros Seleccionados"
   ClientHeight    =   2505
   ClientLeft      =   4245
   ClientTop       =   1380
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FHELPG.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2505
   ScaleWidth      =   5235
   Begin Threed.SSCommand bb_escoger 
      Height          =   375
      Left            =   3390
      TabIndex        =   3
      Top             =   2115
      WhatsThisHelpID =   2104
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Escoger"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
   End
   Begin Threed.SSPanel dl_sp 
      Height          =   330
      Left            =   1890
      TabIndex        =   5
      Top             =   2175
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Alignment       =   1
   End
   Begin Threed.SSCommand bb_cancelar 
      Height          =   375
      Left            =   4305
      TabIndex        =   4
      Top             =   2115
      WhatsThisHelpID =   2012
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
   End
   Begin Threed.SSCommand bb_siguiente 
      Height          =   375
      Left            =   930
      TabIndex        =   1
      Top             =   2115
      WhatsThisHelpID =   2115
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Si&guiente"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
   End
   Begin Threed.SSCommand bb_buscar 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   2115
      WhatsThisHelpID =   2102
      Width           =   900
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Buscar"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
   End
   Begin MSGrid.Grid gr_SQL 
      Height          =   2040
      Left            =   15
      TabIndex        =   2
      Top             =   30
      Width           =   5205
      _Version        =   65536
      _ExtentX        =   9182
      _ExtentY        =   3599
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   0
      X2              =   5220
      Y1              =   2085
      Y2              =   2085
   End
End
Attribute VB_Name = "grid_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bb_buscar_Click()
'PMPasoValores SqlConn&, AGBuscar(1).nombre, 0, AGBuscar(1).tipo, AGBuscar(1).valor
'PMPasoValores SqlConn&, AGBuscar(2).nombre, 0, AGBuscar(2).tipo, AGBuscar(2).valor
'PMPasoValores SqlConn&, AGBuscar(3).nombre, 0, AGBuscar(3).tipo, AGBuscar(3).valor
'For i% = 4 To VGPBuscar%
'PMPasoValores SqlConn&, AGBuscar(i%).nombre, 0, AGBuscar(i%).tipo, AGBuscar(i%).valor
'Next i%
'FPrincipal.pnlTransaccionLine.Caption = ""
'If FMTransmitirRPC(SqlConn&, ServerName$, VGBaseDatos$, (dl_sp.Caption), True, "Help") Then
'PMMapeaGrid SqlConn&, gr_SQL, False
'PMChequea SqlConn&
'If (dl_sp.Caption) <> "sp_hp_catalogo" Then
'If gr_SQL.Rows > MaximoRows% Then
'bb_siguiente.Enabled = True
'Else
'bb_siguiente.Enabled = False
'End If
'Else
'bb_siguiente.Enabled = False
'End If
'End If
End Sub

Private Sub bb_cancelar_Click()
    ReDim Temporales(10)
    PMLimpiaG gr_SQL
    grid_valores.Hide
End Sub

Private Sub bb_escoger_Click()
    gr_SQL_DblClick
End Sub

Private Sub bb_siguiente_Click()
    If gr_SQL.Rows > MaximoRows% Then
'PMPasoValores SqlConn&, AGBuscar(1).nombre, 0, AGBuscar(1).tipo, AGBuscar(1).valor
'PMPasoValores SqlConn&, AGBuscar(2).nombre, 0, AGBuscar(2).tipo, AGBuscar(2).valor
'PMPasoValores SqlConn&, AGBuscar(3).nombre, 0, AGBuscar(3).tipo, Str$(Val(AGBuscar(3).valor) + 1)
'For i% = 4 To VGPBuscar%
'PMPasoValores SqlConn&, AGBuscar(i%).nombre, 0, AGBuscar(i%).tipo, AGBuscar(i%).valor
'Next i%
'gr_SQL.Row = gr_SQL.Rows - 1
'For i% = 1 To VGPSiguiente%
'gr_SQL.Col = AGSiguiente(i%).Col
'PMPasoValores SqlConn&, AGSiguiente(i%).nombre, 0, AGSiguiente(i%).tipo, (gr_SQL.Text)
'Next i%
'FPrincipal.pnlTransaccionLine.Caption = ""
'If FMTransmitirRPC(SqlConn&, ServerName$, VGBaseDatos$, (dl_sp.Caption), True, "Help - Siguiente") Then
'PMMapeaGrid SqlConn&, gr_SQL, True
'PMChequea SqlConn&
'If Val(gr_SQL.Tag) > MaximoRows% Then
'bb_siguiente.Enabled = True
'Else
'bb_siguiente.Enabled = False
'End If
'If (gr_SQL.Rows > (MaximoRows% + 1)) Then
'gr_SQL.TopRow = gr_SQL.Rows - MaximoRows%
'End If
'End If
    End If
End Sub

Private Sub gr_SQL_Click()
    PMLineaG gr_SQL
End Sub

Private Sub gr_SQL_DblClick()
    If Val(gr_SQL.Tag) > 0 Then
        ReDim Temporales(gr_SQL.Cols - 1)
        For c% = 1 To gr_SQL.Cols - 1
            gr_SQL.Col = c%
            Temporales(c%) = gr_SQL.Text
        Next c%
        PMLimpiaG gr_SQL
    Else
        ReDim Temporales(10)
    End If
    grid_valores.Hide
End Sub

Private Sub gr_SQL_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    ReDim Temporales(gr_SQL.Cols)
    '    For c% = 1 To gr_SQL.Cols - 1
    '        gr_SQL.Col = c%
    '        Temporales(c%) = gr_SQL.Text
    '    Next c%
    '    PMLimpiaG gr_SQL
    '    grid_valores.Hide
    'Else
        If KeyAscii = KEY_ESCAPE Then
            ReDim Temporales(10)
            PMLimpiaG gr_SQL
            grid_valores.Hide
        End If
    'End If

End Sub

