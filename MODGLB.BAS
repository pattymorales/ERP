Attribute VB_Name = "MODGLB"
Global Const F5% = 116          'Codigo de la tecla F5
Global Const INSERTAR% = 0      'Codigo de operacion de insercion
Global Const ACTUALIZAR% = 1    'Codigo de operacion de actualizacion
Global Const ELIMINAR% = 2      'Codigo de operacion de eliminacion
Global Const Caracter% = 9      'Caracter tabulador en las listas
Global Const MaximoRows% = 19  'Maximo numero de filas retornadas
Global VGPassword$              'VGPassword del funcionario que opera la aplicación
Global VGProducto$              'Código del producto CARTERA
Global VGFilial$                'Código de la filial de Conexión
Global VGOficina$               'Código de la oficina de Conexión
Global VGRol As String          'Código del rol del usuario
Global VGBusqueda() As String   'Resultado de una búsqueda
Global VGPersonalizacion%       'Chequeo de Personalizacion
Global VGDatabaseName$
Global VGHostName$
Global VGUsuario$               'Codigo de usuario
Global VGUsuarioNombre$         'Nombre del Usuario
Global VGLogTransacciones$      'Archivo de log local del sistema
Global VGBandera$                 'Bandera de ingreso de password de terminal bloqueado
Global VGPaso As Integer
Global VGPasoEstado As Integer
Global VGOperacion As String
Global VGTeclaAyuda As Integer
Global Operacion As Integer
Global Valores()  As String

Type CatalogoUsuario
    Codigo As String
    descripcion As String
End Type

Global VGACatalogo As CatalogoUsuario 'Variable utilizada para llama al catálogo por F5
Global VGTIMERMAIL As Integer        'Contador del poleo del correo elctrónico
Global VGCorreo As Integer           'Permite habilitar el correo electrónico
Global Const KEY_RETURN = &HD
Global Const KEY_ESCAPE = &H1B
Global VGNivelMaximo As String
Global VGPais As String
Global VGBanco As String
Global Const SQLTEXT% = &H23
Global Const SQLARRAY% = &H24
Global Const SQLVARBINARY% = &H25
Global Const SQLINTN% = &H26
Global Const SQLVARCHAR& = &H27
Global Const SQLBINARY% = &H2D
Global Const SQLIMAGE% = &H22
Global Const SQLCHAR& = &H2F
Global Const SQLINT1& = &H30
Global Const SQLBIT% = &H32
Global Const SQLINT2& = &H34
Global Const SQLINT4& = &H38
Global Const SQLMONEY& = &H3C
Global Const SQLDATETIME& = &H3D
Global Const SQLFLT8& = &H3E
Global Const SQLFLTN% = &H6D
Global Const SQLFLT4% = &H3B
Global Const SQLMONEYN% = &H6E
Global Const SQLDATETIMN% = &H6F
Global Const SQLAOPCNT% = &H4B
Global Const SQLAOPSUM% = &H4D
Global Const SQLAOPAVG% = &H4F
Global Const SQLAOPMIN% = &H51
Global Const SQLAOPMAX% = &H52
Global Const SQLAOPANY% = &H53
Global Const SQLAOPNOOP% = &H56
Global Const SQLMONEY4% = &H7A
Global Const SQLDATETIM4% = &H3A

'Variables Globales
Global VTAuxi As Integer
Global Const GFSR_SYSTEMRESOURCES = &H0

'Estandarización en uso de formato de fecha para consultas de fechas
Global VGFecha_SP As String
Global VGFecha_Pref As String
Global Formato_Fecha()  As String
Global VGFechaProceso As String
Global Const CGFormatoBase$ = "mm/dd/yyyy"

'
'CRA24abr9802 DI : Doble declaración de las variables, en el Módulo MAP.BAS y eliminación
'                  de las declaraciones de:Function GetFreeSpace, Function GetSystemDirectory


Function FMValidaTipoDato(TipoDato As String, Valor As Integer)
    FMValidaTipoDato = Valor%
    Select Case TipoDato$
    Case "N"
        If (Valor% <> 8) And ((Valor% < 48) Or (Valor% > 57)) Then
            FMValidaTipoDato = 0
        End If
    Case "A"
        If (Valor% <> 8) And ((Valor% < 65) Or (Valor% > 90)) And ((Valor% < 97) Or (Valor% > 122)) Then
            FMValidaTipoDato = 0
        Else
            FMValidaTipoDato = Asc(UCase$(Chr$(Valor%)))
        End If
    Case "B"
        FMValidaTipoDato = Asc(UCase$(Chr$(Valor%)))
    Case "M"
        If (Valor% <> 8) And (Valor% <> 46) And ((Valor% < 48) Or (Valor% > 57)) Then
            FMValidaTipoDato = 0
        End If
    End Select
End Function

Sub PMLimpiaGrid(grdGrid As Control)
    grdGrid.Rows = 2
    grdGrid.Cols = 2
    For i% = 0 To 1
        grdGrid.Col = i%
        For j% = 0 To 1
            grdGrid.Row = j%
            grdGrid.Text = ""
        Next j%
    Next i%
End Sub

'Función para cargar el formato de la fecha
'PES 1999/07/30
Sub PMCargar_FechaSP(Formato As String)
    For i% = 1 To UBound(Formato_Fecha)
        If Formato_Fecha(i%, 1) = Formato$ Then
           VGFecha_SP$ = Formato_Fecha(i%, 2)
           VGFecha_Pref$ = Formato$
        End If
    Next
End Sub
