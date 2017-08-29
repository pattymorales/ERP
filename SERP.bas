Attribute VB_Name = "SERP"
Function enablemenus(Enabled As Boolean, User As String) As Integer
        If Enabled Then
                If User = "adminnomina" Then
                        FPrincipal.mnuLogon.Enabled = False
                        FPrincipal.mnuLogoff.Enabled = True
                        FPrincipal.mnuPasswd.Enabled = True
                        FPrincipal.mnuUsuarios.Enabled = False
                        FPrincipal.mnuArticulos.Enabled = False
                        FPrincipal.mnuVentas.Enabled = False
                        FPrincipal.mnudatos.Enabled = False
                        FPrincipal.mnuFacturacion.Enabled = False
                Else
                        FPrincipal.mnuLogon.Enabled = False
                        FPrincipal.mnuLogoff.Enabled = True
                        FPrincipal.mnuPasswd.Enabled = True
                        FPrincipal.mnuUsuarios.Enabled = True
                        FPrincipal.mnuArticulos.Enabled = True
                        FPrincipal.mnuVentas.Enabled = True
                        FPrincipal.mnudatos.Enabled = True
                        FPrincipal.mnuFacturacion.Enabled = True
                End If
        Else
                FPrincipal.mnuLogon.Enabled = True
                FPrincipal.mnuLogoff.Enabled = False
                FPrincipal.mnuPasswd.Enabled = False
                FPrincipal.mnuUsuarios.Enabled = False
                FPrincipal.mnuArticulos.Enabled = False
                FPrincipal.mnuVentas.Enabled = False
                FPrincipal.mnudatos.Enabled = False
                FPrincipal.mnuFacturacion.Enabled = False
        End If
        enablemenus = True
End Function
