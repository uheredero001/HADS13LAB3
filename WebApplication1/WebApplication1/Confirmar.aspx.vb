Public Class Confirmar
    Inherits System.Web.UI.Page
    Dim datos As New accesodatosSQL

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        datos.conectar()
        Label2.Text = datos.confirmarRegistro(Request.QueryString("email"), Request.QueryString("numconf"))
        datos.cerrarconexion()
    End Sub

End Class