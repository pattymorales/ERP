Attribute VB_Name = "Inicio"
'*********************************************************
' DESCRIPCION:  Este módulo contiene las rutinas que nos
'               permiten manejar las preferencias.
'*********************************************************
' DECLARACIONES GLOBALES:
' Registro para almacenar los tokens y el valor
'*********************************************************
Type RegistroTOK
    Token As String
    Valor As String
End Type

'***********************************************
' CONSTANTES
'***********************************************
Global ARCHIVOINI$
'***********************************************
' Variables Globales para manejo de *.ini
'***********************************************

Global Preferencias() As String
Global Seccion() As RegistroTOK
Global SMV As Double          'Salario Mínimo Vital
Global NSMV As Double         'Máximo número de SMV
Global VGComp As String
Global VGNComp As String
Global ServerCon As String
Global VGPais As String
Global VGBanco As String

Function Abrir_Archivo(Filename As String) As Integer
'*********************************************************
'Objetivo:  Abre un archivo y devuelve el FileHandler asig-
'           nado a ese archivo
'Input   :  Filename        nombre del archivo
'Output  :  Abrir_Archivo   FileHandler del archivo abierto
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Ago/93  M.Davila            Emisión Inicial
'*********************************************************
    On Error GoTo Salir
    FNum = FreeFile
    Open Filename For Input As #FNum
    Abrir_Archivo = FNum
    Exit Function
Salir:
    Abrir_Archivo = 0
    Exit Function
End Function

Function Buscar_Token(FileNum As Integer, Token As String) As String
'*********************************************************
'Objetivo:  Busca un Token en el archivo .ini abierto y re-
'           torna el valor asociado
'Input   :  FileNum         FileHandler del archivo
'           Token           Token a buscar
'Output  :  Buscar_Token    Valor del token en el archivo
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'01/Ago/93  M.Davila            Emisión Inicial
'*********************************************************
    Dim VTLinea As String
    Seek #FileNum, 1
    Do Until EOF(FileNum)
        Line Input #FileNum, VTLinea
        If InStr(1, VTLinea, "'") = 0 Then
            If InStr(1, VTLinea, Token) > 0 Then
                pos1% = InStr(1, VTLinea, "=")
                If pos1% > 0 Then
                    Buscar_Token = Mid$(VTLinea, pos1% + 1, Len(VTLinea))
                    Exit Function
                End If
            End If
        End If
    Loop
    Buscar_Token = ""
End Function

Sub Escribir_ini(Filename As String)
'*********************************************************
'Objetivo:  Dado el nombre del archivo .ini escribe el
'           contenido del Arreglo Preferencias() en
'           el archivo .ini en forma TOKEN=Valor
'Input   :  FileName        Nombre del Archivo
'Output  :  ninguno
'*********************************************************
    Dim FNum As Integer
    Dim Linea As String
    FNum = FreeFile
    Open Filename For Output As #FNum
    If FNum > 0 Then
            Linea = "[PREFERENCIAS]"
            Print #FNum, Linea
        For i% = 1 To UBound(Preferencias, 1)
            Linea = Preferencias(i%, 1) + "=" + Preferencias(i%, 2)
            Print #FNum, Linea
        Next i%
        Close (FNum)
    End If
End Sub

Function Get_Preferencia(Token As String) As String
'*********************************************************
'Objetivo:  Dado un Token busca el valor actual del mismo
'           en el arreglo Preferencias()
'Input   :  Token               Token del cual se busca el valor
'Output  :  Get_Preferencia     valor actual para Token
'*********************************************************
    For i% = 1 To UBound(Preferencias, 1)
        If Preferencias(i%, 1) = Token Then
            Get_Preferencia = Preferencias(i%, 2)
            Exit Function
        End If
    Next i%
    Get_Preferencia = ""
End Function

Sub Iniciar_Preferencias(Filename As String)
'*********************************************************
'Objetivo:  Dado un nombre de Archivo, copia el contenido
'           del archivo en el arreglo Preferencias()
'Input   :  FileName            Nombre del archivo .ini
'Output  :  ninguno
'*********************************************************
    Dim FNum As Integer
    'Hay que dimensionar este arreglo exactamente al número
    'de preferencias que van a ser manejadas
    ReDim Preferencias(6, 2)
    Preferencias(1, 1) = "USUARIO"
    Preferencias(2, 1) = "FORMATO-FECHA"
    Preferencias(3, 1) = "ODBC"
    Preferencias(4, 1) = "PATH-EXCEL"
    Preferencias(5, 1) = "ARCHIVO_NOM"
    Preferencias(6, 1) = "ARCHIVO_CONTA"
    FNum = Abrir_Archivo(Filename)
    If FNum > 0 Then
        Preferencias(1, 2) = Buscar_Token(FNum, "USUARIO")
        Preferencias(2, 2) = Buscar_Token(FNum, "FORMATO-FECHA")
        If Preferencias(2, 2) = "" Then
            Preferencias(2, 2) = "yyyy/mm/dd"
        End If
        Preferencias(3, 2) = Buscar_Token(FNum, "ODBC")
        Preferencias(4, 2) = Buscar_Token(FNum, "PATH-EXCEL")
        Preferencias(5, 2) = Buscar_Token(FNum, "ARCHIVO_NOM")
        Preferencias(6, 2) = Buscar_Token(FNum, "ARCHIVO_CONTA")
        Close #FNum
    End If
End Sub

Sub Leer_Seccion(FileNum As Integer, posicion As Long)
'*********************************************************
'Objetivo:  Dado un FileHandler y una posicion en el archi-
'           vo donde inicia una sección, lee todos el
'           conjunto TOKEN=Valor dentro esa sección
'Input   :  FileNum        FileHandler del archivo
'           posicion       posición desde donde empieza
'                          la lectura del archivo
'Output  :  ninguno
'*********************************************************
    Seek #FileNum, posicion
    Do Until EOF(FileNum)
        Line Input #FileNum, VTLinea
        If InStr(1, VTLinea, "'") = 0 Then
            If InStr(1, VTLinea, "[") = 0 Then
                pos1% = InStr(1, VTLinea, "=")
                If pos1% > 0 Then
                    Seccion(i%).Token = Mid$(VTLinea, 1, pos1% - 1)
                    Seccion(i%).Valor = Mid$(VTLinea, pos1% + 1, Len(VTLinea))
                    i% = i% + 1
                End If
            Else
                Exit Sub
            End If
        End If
    Loop
End Sub

Function Set_Seccion(FileNum As Integer, Seccion As String) As Long
'*********************************************************
'Objetivo:  Dado un FileHandler y el nombre de una seccion
'           retorna la posición en el archivo dónde inicia
'           la sección.
'Input   :  FileNum         FileHandler del archivo
'           Seccion         String con el nombre de la sección
'Output  :  Set_Seccion     posición de inicio de sección
'*********************************************************
    Dim VTLinea As String
    Seek #FileNum, 1
    Do Until EOF(FileNum)
        Line Input #FileNum, VTLinea
        If InStr(1, VTLinea, "'") = 0 Then
            If InStr(1, VTLinea, "[" + Seccion + "]") > 0 Then
                Set_Seccion = Seek(FileNum)
                Exit Function
            End If
        End If
    Loop
    Set_Seccion = -1
End Function
