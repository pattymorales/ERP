Attribute VB_Name = "MFORMATO"
Function FMDateDiff(Intervalo As String, Fecha1 As String, Fecha2 As String, Formato As String) As Long
'*********************************************************
'Objetivo:  Dados dos strings tipo fecha, calcula la dife-
'           rencia en a�os entre las dos fechas
'           Fecha1 debe ser menor a Fecha2
'           Los formatos soportados son:
'               mm/dd/yy  ;  mm/dd/yyyy
'               dd/mm/yy  ;  dd/mm/yyyy
'               yy/mm/dd  ;  yyyy/mm/dd
'Input:     Fecha1      string con la primera fecha
'           Fecha2      string con la segunda fecha
'           Formato     string con el formato que tienen las fechas
'           Intervalo   string que indica el intervalo
'                       de tiempo entre las dos fechas
'                       "y" = a�os (completos)
'                       "m" = meses
'                       "d" = dias
'Output:    FMDateDiff  intervalo entre las dos fechas.
'                       > 0 si Fecha1 < Fecha2
'                       < 0 si Fecha1 > Fecha2
'                       = 0 si Fecha1 = Fecha2 o Error
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'19/Abr/94  M.Davila            Emisi�n Inicial
'*********************************************************

Dim VTm1$, VTd1$, VTa1$        'para capturar mes, dia y a�o de fecha1
Dim VTm2$, VTd2$, VTa2$        'para capturar mes, dia y a�o de fecha2
Dim VTp11%, VTp12%             'posicion de los dos "/"
Dim VTp21%, VTp22%             'posicion de los dos "/"
Dim VTdd&, VTmm&, VTaa&        'intervalos en dias, meses y a�os

    'Verificar que las fechas sean correctas
    VT% = FMVerFormato(Fecha1, Formato)
    If Not VT% Then
        FMDateDiff& = 0
        Exit Function
    End If
    VT% = FMVerFormato(Fecha2, Formato)
    If Not VT% Then
        FMDateDiff& = 0
        Exit Function
    End If
    
    VTp11% = InStr(1, Fecha1, "/")
    If VTp11% = 0 Then
        FMDateDiff& = 0
        Exit Function
    End If
    VTp12% = InStr(VTp11% + 1, Fecha1, "/")
    If VTp12% = 0 Then
        FMDateDiff& = 0
        Exit Function
    End If
    VTp21% = InStr(1, Fecha2, "/")
    If VTp21% = 0 Then
        FMDateDiff& = 0
        Exit Function
    End If
    VTp22% = InStr(VTp21% + 1, Fecha2, "/")
    If VTp22% = 0 Then
        FMDateDiff& = 0
        Exit Function
    End If
    'extraigo los substrings de mes dia y a�o
    Select Case Formato
    Case "mm/dd/yy"
        VTm1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTd1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTa1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTa1$ = "19" + VTa1$
        VTm2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTd2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTa2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
        VTa2$ = "19" + VTa2$
    Case "mm/dd/yyyy"
        VTm1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTd1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTa1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTm2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTd2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTa2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
    Case "dd/mm/yy"
        VTd1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTm1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTa1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTa1$ = "19" + VTa1$
        VTd2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTm2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTa2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
        VTa2$ = "19" + VTa2$
    Case "dd/mm/yyyy"
        VTd1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTm1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTa1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTd2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTm2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTa2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
    Case "yy/mm/dd"
        VTa1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTm1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTd1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTa1$ = "19" + VTa1$
        VTa2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTm2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTd2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
        VTa2$ = "19" + VTa2$
    Case "yyyy/mm/dd"
        VTa1$ = Mid$(Fecha1, 1, VTp11% - 1)
        VTm1$ = Mid$(Fecha1, VTp11% + 1, VTp12% - VTp11% - 1)
        VTd1$ = Mid$(Fecha1, VTp12% + 1, Len(Fecha1))
        VTa2$ = Mid$(Fecha2, 1, VTp21% - 1)
        VTm2$ = Mid$(Fecha2, VTp21% + 1, VTp22% - VTp21% - 1)
        VTd2$ = Mid$(Fecha2, VTp22% + 1, Len(Fecha2))
    Case Else
        MsgBox "Formato de Fecha " + Formato + " no soportado", 16, "Control Ingreso de Datos"
        FMDateDiff& = 0
        Exit Function
    End Select
    ' ** ENCONTRAR EL INTERVALO ENTRE LAS DOS FECHAS
    Select Case Intervalo
    Case "y"
        'encontrar cual es la fecha mayor
        If Val(VTa1$) < Val(VTa2$) Then   'Fecha1 < Fecha2
            'comparar los meses
            If Val(VTm1$) < Val(VTm2$) Then
                FMDateDiff& = Val(VTa2$) - Val(VTa1$)
            Else
                If Val(VTm1$) > Val(VTm2$) Then
                    FMDateDiff& = Val(VTa2$) - Val(VTa1$) - 1
                Else
                    'meses iguales, comparar dias
                    If Val(VTd1$) <= Val(VTd2$) Then
                        FMDateDiff& = Val(VTa2$) - Val(VTa1$)
                    Else
                        FMDateDiff& = Val(VTa2$) - Val(VTa1$) - 1
                    End If
                End If
            End If
        Else
            If Val(VTa1$) > Val(VTa2$) Then  'Fecha1 > Fecha2
                'comparar los meses
                If Val(VTm2$) < Val(VTm1$) Then
                    FMDateDiff& = -(Val(VTa1$) - Val(VTa2$))
                Else
                    If Val(VTm2$) > Val(VTm1$) Then
                        FMDateDiff& = -(Val(VTa1$) - Val(VTa2$) - 1)
                    Else
                        'meses iguales, comparar dias
                        If Val(VTd2$) <= Val(VTd1$) Then
                            FMDateDiff& = -(Val(VTa1$) - Val(VTa2$))
                        Else
                            FMDateDiff& = -(Val(VTa1$) - Val(VTa2$) - 1)
                        End If
                    End If
                End If
            Else
                'son fechas del mismo a�o
                FMDateDiff& = 0
            End If
        End If
    Case "m"
        'encontrar cual fecha es mayor
        If Val(VTa1$) < Val(VTa2$) Then   'Fecha1 < Fecha2
            'sumar los meses de a�os enteros
            VTmm& = (Val(VTa2$) - Val(VTa1$) - 1) * 12
            'sumar los meses del a�o en la Fecha1
            VTmm& = VTmm& + (12 - Val(VTm1$))
            'sumar los meses del a�o en la Fecha 2
            VTmm& = VTmm& + (Val(VTm2$) - 1)
            If Val(VTd1$) <= Val(VTd2$) Then
                VTmm& = VTmm& + 1
            End If
            FMDateDiff& = VTmm&
        Else
            If Val(VTa1$) > Val(VTa2$) Then  'Fecha1 > Fecha2
                'sumar los meses de a�os enteros
                VTmm& = (Val(VTa1$) - Val(VTa2$) - 1) * 12
                'sumar los meses del a�o en la Fecha2
                VTmm& = VTmm& + (12 - Val(VTm2$))
                'sumar los meses del a�o en la Fecha1
                VTmm& = VTmm& + (Val(VTm1$) - 1)
                If Val(VTd2$) <= Val(VTd1$) Then
                    VTmm& = VTmm& + 1
                End If
                FMDateDiff& = -VTmm&
            Else
                'son fechas del mismo a�o
                If Val(VTm1$) < Val(VTm2$) Then
                    'sumar meses enteros
                    VTmm& = Val(VTm2$) - Val(VTm1$) - 1
                    If Val(VTd1$) <= Val(VTd2$) Then
                        VTmm& = VTmm& + 1
                    End If
                    FMDateDiff& = VTmm&
                Else
                    If Val(VTm1$) > Val(VTm2$) Then
                        'sumar meses enteros
                        VTmm& = Val(VTm1$) - Val(VTm2$) - 1
                        If Val(VTd2$) <= Val(VTd1$) Then
                            VTmm& = VTmm& + 1
                        End If
                        FMDateDiff& = -VTmm&
                    Else
                        FMDateDiff& = 0
                    End If
                End If
            End If
        End If
    Case "d"
        VTdd& = 0
        'calcular la fecha mayor
        If Val(VTa1$) < Val(VTa2$) Then 'Fecha1 < Fecha2
            'calcular dias de a�os completos
            For i% = (Val(VTa1$) + 1) To (Val(VTa2$) - 1)
                If (i% Mod 4) = 0 Then
                    If (i% Mod 100 = 0) And (i% Mod 400 <> 0) Then
                        VTdd& = VTdd& + 365
                    Else
                        VTdd& = VTdd& + 366
                    End If
                Else
                    VTdd& = VTdd& + 365
                End If
            Next i%
            'calcular los dias en el a�o de Fecha 1
            If (Val(VTa1$) Mod 4) = 0 Then
                If (Val(VTa1$) Mod 100 = 0) And (Val(VTa1$) Mod 400 <> 0) Then
                    vtbis% = False
                Else
                    vtbis% = True
                End If
            Else
                vtbis% = False
            End If
            'en el mes de Fecha1
            Select Case Val(VTm1$)
            Case 1, 3, 5, 7, 8, 10, 12
                VTdd& = VTdd& + (31 - Val(VTd1$))
            Case 4, 6, 9, 11
                VTdd& = VTdd& + (30 - Val(VTd1$))
            Case 2
                If vtbis% Then
                    VTdd& = VTdd& + (29 - Val(VTd1$))
                Else
                    VTdd& = VTdd& + (28 - Val(VTd1$))
                End If
            End Select
            'el resto de dias
            For i% = Val(VTm1$) + 1 To 12
                Select Case i%
                Case 1, 3, 5, 7, 8, 10, 12
                    VTdd& = VTdd& + 31
                Case 4, 6, 9, 11
                    VTdd& = VTdd& + 30
                Case 2
                    If vtbis% Then
                        VTdd& = VTdd& + 29
                    Else
                        VTdd& = VTdd& + 28
                    End If
                End Select
            Next i%
            'calcular los dias en el a�o de Fecha 2
            If (Val(VTa2$) Mod 4) = 0 Then
                If (Val(VTa2$) Mod 100 = 0) And (Val(VTa2$) Mod 400 <> 0) Then
                    vtbis% = False
                Else
                    vtbis% = True
                End If
            Else
                vtbis% = False
            End If
            'en el mes de Fecha2
            VTdd& = VTdd& + Val(VTd2$)
            'en los meses anteriores
            For i% = 1 To Val(VTm2$) - 1
                Select Case i%
                Case 1, 3, 5, 7, 8, 10, 12
                    VTdd& = VTdd& + 31
                Case 4, 6, 9, 11
                    VTdd& = VTdd& + 30
                Case 2
                    If vtbis% Then
                        VTdd& = VTdd& + 29
                    Else
                        VTdd& = VTdd& + 28
                    End If
                End Select
            Next i%
            FMDateDiff& = VTdd&
        Else
            If Val(VTa1$) > Val(VTa2$) Then 'Fecha1 > Fecha2
                'calcular dias de a�os completos
                For i% = (Val(VTa2$) + 1) To (Val(VTa1$) - 1)
                    If (i% Mod 4) = 0 Then
                        If (i% Mod 100 = 0) And (i% Mod 400 <> 0) Then
                            VTdd& = VTdd& + 365
                        Else
                            VTdd& = VTdd& + 366
                        End If
                    Else
                        VTdd& = VTdd& + 365
                    End If
                Next i%
                'calcular los dias en el a�o de Fecha2
                If (Val(VTa2$) Mod 4) = 0 Then
                    If (Val(VTa2$) Mod 100 = 0) And (Val(VTa2$) Mod 400 <> 0) Then
                        vtbis% = False
                    Else
                        vtbis% = True
                    End If
                Else
                    vtbis% = False
                End If
                'en el mes de Fecha2
                Select Case Val(VTm2$)
                Case 1, 3, 5, 7, 8, 10, 12
                    VTdd& = VTdd& + (31 - Val(VTd2$))
                Case 4, 6, 9, 11
                    VTdd& = VTdd& + (30 - Val(VTd2$))
                Case 2
                    If vtbis% Then
                        VTdd& = VTdd& + (29 - Val(VTd2$))
                    Else
                        VTdd& = VTdd& + (28 - Val(VTd2$))
                    End If
                End Select
                'el resto de dias
                For i% = Val(VTm2$) + 1 To 12
                    Select Case i%
                    Case 1, 3, 5, 7, 8, 10, 12
                        VTdd& = VTdd& + 31
                    Case 4, 6, 9, 11
                        VTdd& = VTdd& + 30
                    Case 2
                        If vtbis% Then
                            VTdd& = VTdd& + 29
                        Else
                            VTdd& = VTdd& + 28
                        End If
                    End Select
                Next i%
                'calcular los dias en el a�o de Fecha1
                If (Val(VTa1$) Mod 4) = 0 Then
                    If (Val(VTa1$) Mod 100 = 0) And (Val(VTa1$) Mod 400 <> 0) Then
                        vtbis% = False
                    Else
                        vtbis% = True
                    End If
                Else
                    vtbis% = False
                End If
                'en el mes de Fecha1
                VTdd& = VTdd& + Val(VTd1$)
                'en los meses anteriores
                For i% = 1 To Val(VTm1$) - 1
                    Select Case i%
                    Case 1, 3, 5, 7, 8, 10, 12
                        VTdd& = VTdd& + 31
                    Case 4, 6, 9, 11
                        VTdd& = VTdd& + 30
                    Case 2
                        If vtbis% Then
                            VTdd& = VTdd& + 29
                        Else
                            VTdd& = VTdd& + 28
                        End If
                    End Select
                Next i%
                FMDateDiff& = -VTdd&
            Else
                ' a�os iguales
                If (Val(VTa1$) Mod 4) = 0 Then
                    If (Val(VTa1$) Mod 100 = 0) And (Val(VTa1$) Mod 400 <> 0) Then
                        vtbis% = False
                    Else
                        vtbis% = True
                    End If
                Else
                    vtbis% = False
                End If
                If Val(VTm1$) < Val(VTm2$) Then
                    'los dias que faltan del mes de Fecha1
                    Select Case Val(VTm1$)
                    Case 1, 3, 5, 7, 8, 10, 12
                        VTdd& = VTdd& + (31 - Val(VTd1$))
                    Case 4, 6, 9, 11
                        VTdd& = VTdd& + (30 - Val(VTd1$))
                    Case 2
                        If vtbis% Then
                            VTdd& = VTdd& + (29 - Val(VTd1$))
                        Else
                            VTdd& = VTdd& + (28 - Val(VTd1$))
                        End If
                    End Select
                    'los dias de meses completos entre Fecha1 y Fecha2
                    For i% = Val(VTm1$) + 1 To Val(VTm2$) - 1
                        Select Case i%
                        Case 1, 3, 5, 7, 8, 10, 12
                            VTdd& = VTdd& + 31
                        Case 4, 6, 9, 11
                            VTdd& = VTdd& + 30
                        Case 2
                            If vtbis% Then
                                VTdd& = VTdd& + 29
                            Else
                                VTdd& = VTdd& + 28
                            End If
                        End Select
                    Next i%
                    'los dias transcurridos del mes de Fecha2
                    VTdd& = VTdd& + Val(VTd2$)
                    FMDateDiff& = VTdd&
                Else
                    If Val(VTm1$) > Val(VTm2$) Then
                        'los dias que faltan del mes de Fecha2
                        Select Case Val(VTm2$)
                        Case 1, 3, 5, 7, 8, 10, 12
                            VTdd& = VTdd& + (31 - Val(VTd2$))
                        Case 4, 6, 9, 11
                            VTdd& = VTdd& + (30 - Val(VTd2$))
                        Case 2
                            If vtbis% Then
                                VTdd& = VTdd& + (29 - Val(VTd2$))
                            Else
                                VTdd& = VTdd& + (28 - Val(VTd2$))
                            End If
                        End Select
                        'los dias de meses completos entre Fecha1 y Fecha2
                        For i% = Val(VTm2$) + 1 To Val(VTm1$) - 1
                            Select Case i%
                            Case 1, 3, 5, 7, 8, 10, 12
                                VTdd& = VTdd& + 31
                            Case 4, 6, 9, 11
                                VTdd& = VTdd& + 30
                            Case 2
                                If vtbis% Then
                                    VTdd& = VTdd& + 29
                                Else
                                    VTdd& = VTdd& + 28
                                End If
                            End Select
                        Next i%
                        'los dias transcurridos del mes de Fecha1
                        VTdd& = VTdd& + Val(VTd1$)
                        FMDateDiff& = -VTdd&
                    Else
                        'a�o/mes iguales
                        VTdd& = Val(VTd2$) - Val(VTd1$)
                        FMDateDiff& = VTdd&
                    End If
                End If
            End If
        End If
    End Select
End Function

Function FMMascaraFecha(Formato As String) As String
'*********************************************************
'Objetivo:  Dado el string con el formato retorna la mascara
'           apropiada
'           Los formatos soportados son:
'               mm/dd/yy  ;  mm/dd/yyyy
'               dd/mm/yy  ;  dd/mm/yyyy
'               yy/mm/dd  ;  yyyy/mm/dd
'Input:     Formato     string con el formato de fecha
'           Intervalo   stromg que indica el intervalo
'                       de tiempo entre las dos fechas
'                       "y" = a�os (completos)
'                       "m" = meses
'                       "d" = dias
'Output:    FMMascaraFecha  mascara apropiada
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'22/Abr/94  M.Davila            Emisi�n Inicial
'*********************************************************
    Select Case Formato
    Case "yy/mm/dd", "mm/dd/yy", "dd/mm/yy"
        FMMascaraFecha = "##/##/##"
    Case "mm/dd/yyyy", "dd/mm/yyyy"
        FMMascaraFecha = "##/##/####"
    Case "yyyy/mm/dd"
        FMMascaraFecha = "####/##/##"
    Case Else
        FMMascaraFecha = ""
    End Select
End Function

Function FMVerFormato(Fecha As String, Formato As String) As Integer
'*********************************************************
'Objetivo:  Dado un string que es una fecha, verifica
'           que este en un formato dado.
'           Los formatos soportados son:
'               mm/dd/yy  ;  mm/dd/yyyy
'               dd/mm/yy  ;  dd/mm/yyyy
'               yy/mm/dd  ;  yyyy/mm/dd
'Input:     Fecha       es el string con la fecha
'           Formato     es un string con el formato a verificar
'Output:    FMVerFormato  True/False depende de si Fecha esta
'                           o no en Formato
'*********************************************************
'                    MODIFICACIONES
'FECHA      AUTOR               RAZON
'19/Abr/94  M.Davila            Emisi�n Inicial
'*********************************************************

Dim vtm$, vtd$, vta$        'para capturar mes, dia y a�o
Dim Vtp1%, VTp2%            'posicion de los dos "/"
    Vtp1% = InStr(1, Fecha, "/")
    If Vtp1% = 0 Then
        FMVerFormato% = False
        Exit Function
    End If
    VTp2% = InStr(Vtp1% + 1, Fecha, "/")
    If VTp2% = 0 Then
        FMVerFormato% = False
        Exit Function
    End If
    'extraigo los substrings de mes dia y a�o
    Select Case Formato
    Case "mm/dd/yy"
        vtm$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtd$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "mm/dd/yyyy"
        vtm$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtd$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case "dd/mm/yy"
        vtd$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "dd/mm/yyyy"
        vtd$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case "yy/mm/dd"
        vta$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vtd$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "yyyy/mm/dd"
        vta$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vtd$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case Else
        MsgBox "Formato de Fecha " + Formato + " no soportado", 16, "Control Ingreso de Datos"
        FMVerFormato% = False
        Exit Function
    End Select
    'verificacion de la validez de mes dia y a�o
    If Val(vta$) < 1753 Or Val(vta$) > 9999 Then
        FMVerFormato% = False
        Exit Function
    End If
    If Val(vtm$) < 1 Or Val(vtm$) > 12 Then
        FMVerFormato% = False
        Exit Function
    End If
    If Val(vtd$) < 1 Then
        FMVerFormato% = False
        Exit Function
    End If
    Select Case Val(vtm$)
    Case 1, 3, 5, 7, 8, 10, 12
        If Val(vtd$) > 31 Then
            FMVerFormato% = False
            Exit Function
        End If
    Case 4, 6, 9, 11
        If Val(vtd$) > 30 Then
            FMVerFormato% = False
            Exit Function
        End If
    Case 2
        If (Val(vta$) Mod 4) = 0 Then
            If ((Val(vta$) Mod 100) = 0) And ((Val(vta$) Mod 400) <> 0) Then
                If Val(vtd$) > 28 Then
                    FMVerFormato% = False
                    Exit Function
                End If
            Else
                If Val(vtd$) > 29 Then
                    FMVerFormato% = False
                    Exit Function
                End If
            End If
        Else
            If Val(vtd$) > 28 Then
                FMVerFormato% = False
                Exit Function
            End If
        End If
    End Select
    FMVerFormato% = True
End Function

Function FMConvFecha(Fecha As String, FormatoA As String, FormatoB As String) As String
'*********************************************************
'Objetivo:  Dado un string que es una fecha, cambia el formato
'           de fecha de FormatoA a FormatoB
'           Formatos soportados:
'               mm/dd/yy  ;  mm/dd/yyyy
'               dd/mm/yy  ;  dd/mm/yyyy
'               yy/mm/dd  ;  yyyy/mm/dd
'Input:     Fecha       es el string con la fecha
'           FormatoA    string con formato original
'           FormatoB    string con formato destino
'Output:    FMConvFecha string convertido de formato
'                       si no hay forma de convertir retorna ""
'*********************************************************

Dim vta$, vtm$, vtd$    'substring con a�o, mes, d�a
Dim Vtp1%, VTp2%        'posiciones de "/"
Dim VTFecha$            'auxiliar para la fecha

    'verificamos que Fecha de acuerdo a FormatoA sea valida
    VT% = FMVerFormato(Fecha, FormatoA)
    If Not VT% Then
        FMConvFecha$ = ""
        Exit Function
    End If
    'extraigo los valores de a�o, mes, dia de la fecha
    Vtp1% = InStr(1, Fecha, "/")
    If Vtp1% = 0 Then
        FMConvFecha$ = ""
        Exit Function
    End If
    VTp2% = InStr(Vtp1% + 1, Fecha, "/")
    If VTp2% = 0 Then
        FMConvFecha$ = ""
        Exit Function
    End If
    Select Case FormatoA
    Case "mm/dd/yy"
        vtm$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtd$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "mm/dd/yyyy"
        vtm$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtd$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case "dd/mm/yy"
        vtd$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "dd/mm/yyyy"
        vtd$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vta$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case "yy/mm/dd"
        vta$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vtd$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
        vta$ = "19" + vta$
    Case "yyyy/mm/dd"
        vta$ = Mid$(Fecha, 1, Vtp1% - 1)
        vtm$ = Mid$(Fecha, Vtp1% + 1, VTp2% - Vtp1% - 1)
        vtd$ = Mid$(Fecha, VTp2% + 1, Len(Fecha))
    Case Else
        MsgBox "Formato Origen: " + Formato + " no soportado", 16, "Control Ingreso de Datos"
        FMConvFecha = ""
        Exit Function
    End Select
    'Convertimos fecha a nuevo formato
    Select Case FormatoB
    Case "mm/dd/yy"
        If Len(vta$) = 4 Then
            vta$ = Mid$(vta$, 3, 2)
        End If
        VTFecha$ = vtm$ & "/" & vtd$ & "/" & vta$
    Case "mm/dd/yyyy"
        If Len(vta$) = 2 Then
            vta$ = "19" + vta$
        End If
        VTFecha$ = vtm$ & "/" & vtd$ & "/" & vta$
    Case "dd/mm/yy"
        If Len(vta$) = 4 Then
            vta$ = Mid$(vta$, 3, 2)
        End If
        VTFecha$ = vtd$ & "/" & vtm$ & "/" & vta$
    Case "dd/mm/yyyy"
        If Len(vta$) = 2 Then
            vta$ = "19" + vta$
        End If
        VTFecha$ = vtd$ & "/" & vtm$ & "/" & vta$
    Case "yy/mm/dd"
        If Len(vta$) = 4 Then
            vta$ = Mid$(vta$, 3, 2)
        End If
        VTFecha$ = vta$ & "/" & vtm$ & "/" & vtd$
    Case "yyyy/mm/dd"
        If Len(vta$) = 2 Then
            vta$ = "19" + vta$
        End If
        VTFecha$ = vta$ & "/" & vtm$ & "/" & vtd$
    Case Else
        MsgBox "Formato Destino: " + Formato + " no soportado", 16, "Control Ingreso de Datos"
        FMConvFecha = ""
        Exit Function
    End Select
    FMConvFecha$ = VTFecha$
End Function

Function FMFormatoFecha(flocal As String)
    Select Case flocal
    Case "mm/dd/yyyy"
        FMFormatoFecha = 101
    Case "dd/mm/yyyy"
        FMFormatoFecha = 103
    Case "yyyy/mm/dd"
        FMFormatoFecha = 111
    End Select
End Function
