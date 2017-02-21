Public Class CambiarPassword
    Inherits System.Web.UI.Page
    Private Shared datos As accesodatosSQL

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        datos.conectar()
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        datos.cerrarconexion()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As String

        result = datos.existeUserConPregunta(TextBox1.Text, DropDownList1.Text, TextBox5.Text)
        If (result.Equals("Ok")) Then
            Label6.Text = datos.cambiaPassword(TextBox1.Text, TextBox4.Text)
        Else
            Label6.Text = "Hay algún dato erroneo"
        End If
    End Sub
End Class