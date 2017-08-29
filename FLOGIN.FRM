VERSION 5.00
Object = "{237F7DF8-8CC4-4DEF-9736-78A40ACD7B87}#9.0#0"; "JMButton.ocx"
Begin VB.Form Flogin 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FLOGIN.frx":0000
   ScaleHeight     =   2460
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmpanel 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   3135
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
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Text            =   "paty"
         Top             =   240
         Width           =   1695
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
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "paty123"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbldatos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbldatos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin JMButton.JMBcontrol JMBaceptar 
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1800
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
      Picture         =   "FLOGIN.frx":12DE
      Caption         =   "Aceptar"
      WordWrap        =   -1  'True
      Border          =   3
   End
   Begin JMButton.JMBcontrol JMBaceptar 
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
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
      Picture         =   "FLOGIN.frx":1BB8
      Caption         =   "Salir"
      CaptionPlace    =   4
      WordWrap        =   -1  'True
      Border          =   3
   End
End
Attribute VB_Name = "Flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim VLOK As Boolean

Private Sub Form_Load()
    VLOK = False
    ARCHIVOINI$ = App.Path + "\erp.ini"
    Iniciar_Preferencias (ARCHIVOINI$)
    txtuser(0).Text = Get_Preferencia("USUARIO")
    VGodbc = Get_Preferencia("ODBC")
    VGFormatoFecha = Get_Preferencia("FORMATO-FECHA")
End Sub

Private Sub JMBaceptar_Click(Index As Integer)
    Select Case Index
        Case 0
            PLVerifica_clave
            Unload Me
            FPrincipal.Show 0
        Case 1
            Unload Me
    End Select
End Sub

Private Sub PLVerifica_clave()
Dim rstLogin As Recordset
Dim strT As String
Dim errObj As Error

Screen.MousePointer = 11
On Error GoTo ErrorNinguno
    If txtuser(0).Text = "adminerp" Then
        Set dbErp = OpenDatabase(VGodbc, False, False, "ODBC;DataBase=master;UID=adminerp;PWD=" & txtuser(1).Text)
        VGLOGIN = "adminerp"
        VGPASSBASE = txtuser(1).Text
        strT = "SELECT name FROM syslogins WHERE name = '" & VGLOGIN & "'"
        Set rstLogin = dbErp.OpenRecordset(strT)
        If rstLogin.EOF Then
            MsgBox "No existe el usuario", 48, "Warning"
            Screen.MousePointer = 0
            VLOK = enablemenus(False, VGLOGIN$)
            Exit Sub
        End If
        dbErp.Close
        Set dbErp = OpenDatabase(VGodbc, False, False, "ODBC;DataBase=db_erp;UID=adminerp;PWD=" & txtuser(1).Text)
        VLOK = enablemenus(True, VGLOGIN$)
    Else
        VGCid = txtuser(0).Text
        Set dbErp = OpenDatabase(VGodbc, False, False, "ODBC;DataBase=db_erp;UID=" & txtuser(0).Text & ";PWD=" & txtuser(1).Text)
        strT = "SELECT us_codigo, us_password, us_estado" & _
        " FROM er_user WHERE us_login = '" & VGCid & "'"
        Set rstLogin = dbErp.OpenRecordset(strT)
        If rstLogin.EOF Then
            MsgBox "No existe el usuario", 48, "Warning"
            Screen.MousePointer = 0
            VLOK = enablemenus(False, VGLOGIN$)
            Exit Sub
        End If
        VGCpass = txtuser(1).Text
        'Si el usuario existe
        If Trim(rstLogin.Fields("us_password")) <> RTrim(VGCpass) Then
            MsgBox "Clave de acceso errada. Favor verifique", vbCritical, "Mensaje..."
            txtuser(1) = ""
            txtuser(1).SetFocus
            Screen.MousePointer = 0
            VLOK = enablemenus(False, VGLOGIN$)
            Exit Sub
        End If
        If rstLogin.Fields("us_estado") = "O" Then
            MsgBox "La clave ha sido bloqueada.", vbCritical, "Mensaje..."
            txtuser(0) = ""
            txtuser(1) = ""
            txtuser(0).SetFocus
            Screen.MousePointer = 0
            VLOK = enablemenus(False, VGLOGIN$)
            Exit Sub
        End If
        VGLOGIN = txtuser(0).Text
        VGPASSBASE = txtuser(1).Text
        VLOK = enablemenus(True, VGLOGIN$)
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrorNinguno:
    VGdisperror = False
    For Each errObj In Errors
        MsgBox errObj.Number & " " & errObj.Description, vbCritical, "Login"
        VGdisperror = True
    Next
    If VGdisperror = False Then
        MsgBox "Error al comparar el password", vbCritical, "Login"
    End If
    Screen.MousePointer = 0
    VLOK = enablemenus(False, VGLOGIN$)
End Sub

Private Sub sscAceptar_Click()
        PLVerifica_clave
        Unload Me
        FPrincipal.Show 0
End Sub

Private Sub sscCancelar_Click()
    Unload Me
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


