VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form FCambioPassword 
   Caption         =   "Cambio  Password"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmpanel 
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4215
      Begin VB.TextBox txtuser 
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
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtuser 
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
         Index           =   1
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtuser 
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
         Index           =   0
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label lbldatos 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lbldatos 
         Caption         =   "Clave Actual:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbldatos 
         Caption         =   "Clave anterior:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   200
         Width           =   1455
      End
   End
   Begin JMButton.JMBcontrol jmbboton 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "FCambioPassword.frx":0000
      Caption         =   "Cambiar"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol jmbboton 
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      Picture         =   "FCambioPassword.frx":08DA
      Caption         =   "Salir"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "FCambioPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub jmbboton_Click(Index As Integer)
    Select Case Index
        Case 0
            PLCambiapassword
        Case 1
            Unload Me
    End Select
End Sub

Private Sub PLCambiapassword()
Dim rstLogin As Recordset
Dim pass_anterior As String
Dim estado As String
    For i% = 0 To 2
        If txtuser(i%).Text = "" Then
            Select Case i%
                Case 0
                    MsgBox "Debe ingresar la antigua clave", vbCritical, "Mensaje ..."
                Case 1
                    MsgBox "Debe ingresar la nueve clave", vbCritical, "Mensaje ..."
                Case 3
                    MsgBox "Debe repetir la nueve clave", vbCritical, "Mensaje ..."
            End Select
            txtuser(i%).SetFocus
            Exit Sub
        End If
    Next i%
On Error GoTo ErrorCambiar
    strT = "SELECT us_password, us_estado FROM er_user WHERE us_login = '" & VGLOGIN & "'"
    Set rstLogin = dbErp.OpenRecordset(strT)
    If Not rstLogin.EOF Then
        pass_anterior = rstLogin(0)
        estado = rstLogin(1)
    Else
        MsgBox "El usuario a eliminar no existe", vbCritical, "Login"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If estado = "I" Then
        MsgBox "El estado del usuario no permite modificar el password", vbCritical, "Login"
        Screen.MousePointer = 0
        Exit Sub
    End If
    If txtuser(1).Text <> txtuser(2).Text Then
        MsgBox "La nueva clave no es correcta", vbCritical, "Mensaje ..."
        txtuser(1).SetFocus
        Exit Sub
    End If
    If Trim$(pass_anterior) <> Trim$(txtuser(0).Text) Then
        MsgBox "La clave antigua no es correcta", vbCritical, "Mensaje ..."
        txtuser(0).SetFocus
        Exit Sub
    End If

    Command_SQLx = "sp_user U, null," & VGLOGIN & ", " & txtuser(1).Text & " , estado, null, S"
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    VGPASSBASE$ = txtuser(1).Text
    Command_SQLx = "sp_password " & pass_anterior & ", " & VGPASSBASE$ & "," & VGLOGIN
    Set rstLogin = dbErp.OpenRecordset(Command_SQLx, dbOpenDynaset, dbSQLPassThrough)
    MsgBox "La clave fue cambiada correctamente", vbOKOnly, "Mensaje ..."
    Exit Sub
ErrorCambiar:
    VGdisperror = False
    For Each errObj In Errors
        Var% = MsgBox(errObj.Number & " " & errObj.Description, vbCritical, "Mensaje ...")
        VGdisperror = True
    Next
    If VGdisperror = False Then
        Var% = MsgBox("Error al cambiar la clave", vbCritical, "Mensaje ...")
    End If
End Sub

Private Sub txtuser_GotFocus(Index As Integer)
    txtuser(Index).SelStart = 0
    txtuser(Index).SelLength = 15
End Sub

Private Sub txtuser_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii% <> 8) Then
        If Len(txtuser(Index).Text) > 14 Then
            KeyAscii% = 0
        End If
    End If
End Sub
