
Public Class Inicio
    Inherits System.Web.UI.Page
    Private Shared datos As New accesodatosSQL

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim result As String
        result = datos.conectar()
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        datos.cerrarconexion()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As String
        result = datos.comprobarLogin(TextBox1.Text, TextBox2.Text)
        If (result.Equals("Ok")) Then
            Session("user") = TextBox1.Text
            Response.Redirect("Aplicacion.aspx")
        Else
            Label3.Text = "Error en login"
        End If

    End Sub
End Class