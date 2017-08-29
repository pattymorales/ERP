Attribute VB_Name = "MHELPG"
Type RegistroREC
    nombre As String
    Valor  As String
    tipo   As Integer
End Type

Type RegistroSIG
    nombre As String
    Col As Integer
    tipo As Integer
End Type

Global AGBuscar() As RegistroREC
Global AGSiguiente() As RegistroSIG
Global Temporales() As String

Global VGPBuscar%
Global VGPSiguiente%
Global VGBaseDatos$

Function FLFormatFecha(Fecha As String) As String
    aa$ = Mid$(Fecha, 1, 4)
    mm$ = Mid$(Fecha, 6, 2)
    dd$ = Mid$(Fecha, 9, 2)
    s1$ = Mid$(Fecha, 5, 1)
    s2$ = Mid$(Fecha, 8, 1)
    If s1$ = "/" Then
       If s2$ = "/" Then
          a% = Val(aa$)
          If a% >= 1753 And a% <= 9999 Then
             m% = Val(mm$)
             If m% >= 1 And m% <= 12 Then
                d% = Val(dd$)
                If d% >= 1 And d% <= 31 Then
                   FLFormatFecha = mm$ + "/" + dd$ + "/" + aa$
                   Exit Function
                End If
             End If
          End If
       End If
    End If
    FLFormatFecha = ""
    Exit Function
End Function

Function FLHora(Hora As String) As String
    hh$ = Mid$(Hora, 1, 2)
    mm$ = Mid$(Hora, 4, 2)
    s1$ = Mid$(Hora, 3, 1)
    If s1$ = ":" Then
          h% = Val(hh$)
          If a% >= 0 And a% <= 23 Then
             m% = Val(mm$)
             If m% >= 0 And m% <= 59 Then
                   If m% < 10 Then
                    FLHora = hh$ + ":" + "0" + LTrim$(Str$(m%))
                   Else
                    FLHora = hh$ + ":" + mm$
                   End If
                   Exit Function
             End If
          End If
    End If
    FLHora = ""
    Exit Function
End Function

Sub MensajesVB(Clase As Integer, Numero As Integer, Texto As String)
    Select Case Clase
    Case 0
        Titulo$ = "Error de Conección"
    Case 1
        Titulo$ = "Control de Ingreso de Valores"
    Case 2
        Titulo$ = "Valores Permisibles"
    Case 3
        Titulo$ = "Navegación de Registros"
    End Select
    Beep
    MsgBox Texto$, 0 + Numero%, Titulo$
End Sub

Sub PasoValores(Row%, Lista As Control, Valores() As String)
    Lista.Row = Row%
    For i% = 1 To Lista.Cols - 1
        Lista.Col = i%
        Valores(i%) = Lista.Text
    Next i%
End Sub

Sub PasoValoresGrid(Row%, Lista As Control, Valores() As String)
    ReDim Valores(Lista.Cols - 1) As String
    Lista.Row = Row%
    For i% = 1 To Lista.Cols - 1
        Lista.Col = i%
        Valores(i%) = Lista.Text
    Next i%
End Sub

Sub PasoValoresLista(Row As Integer, Lista As Control, Valores() As String)
    ReDim Valores(Lista.Cols - 1) As String
    Valores(0) = Persona$
    Lista.Row = Row%
    For i% = 1 To Lista.Cols - 1
        Lista.Col = i%
        Valores(i%) = Lista.Text
    Next i%
End Sub

Sub PMBuscarG(posicion As Integer, nombre As String, Valor As String, tipo As Integer)
    AGBuscar(posicion).nombre = nombre
    AGBuscar(posicion).Valor = Valor
    AGBuscar(posicion).tipo = tipo
End Sub

Sub PMHelpG(base_datos As String, stored_procedure As String, par_buscar As Integer, par_sig As Integer)
    grid_valores.dl_sp.Caption = stored_procedure
    VGPBuscar% = par_buscar
    VGPSiguiente% = par_sig
    VGBaseDatos$ = base_datos
    ReDim AGBuscar(VGPBuscar%)
    ReDim AGSiguiente(VGPSiguiente%)
End Sub

Sub PMLimpiaG(GridControl As Control)
  If TypeOf GridControl Is Grid Then
     GridControl.Tag = "0"
     For i% = 0 To 1
        For j% = 0 To 1
            GridControl.Row = i%
            GridControl.Col = j%
            GridControl.Text = ""
        Next j%
     Next i%
  End If
End Sub

Sub PMLineaG(GridControl As Control)
    If GridControl.Rows >= 2 Then
        If GridControl.Row > 0 Then
            GridControl.Col = 0
            GridControl.SelStartCol = 1
            GridControl.SelEndCol = GridControl.Cols - 1
            GridControl.SelStartRow = GridControl.Row
            GridControl.SelEndRow = GridControl.Row
        End If
    End If
End Sub

Sub PMProcesG()
 If grid_valores.dl_sp.Caption <> "sp_hp_catalogo" Then
     If grid_valores.gr_SQL.Rows > MaximoRows% Then
         grid_valores.bb_siguiente.Enabled = True
     Else
        grid_valores.bb_siguiente.Enabled = False
     End If
 Else
    grid_valores.bb_siguiente.Enabled = False
 End If
End Sub

Sub PMSigteG(posicion As Integer, nombre As String, Col As Integer, tipo As Integer)
    AGSiguiente(posicion).nombre = nombre
    AGSiguiente(posicion).Col = Col
    AGSiguiente(posicion).tipo = tipo
End Sub

Sub PMLlenarGrid(stored_procedure As String)
    StoredProcedure = stored_procedure
    Fgrid_catalogo.Show 1
End Sub
