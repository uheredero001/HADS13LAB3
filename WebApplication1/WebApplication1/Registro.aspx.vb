Public Class Registro
    Inherits System.Web.UI.Page
    Private Shared datos As New accesodatosSQL


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        datos.conectar()
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        datos.cerrarconexion()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim emailSender As New envioEmail
        Dim numConf As Integer
        Randomize()
        numConf = CInt(Rnd() * 9000000) + 1000000
        Respuesta.Text = datos.insertar(TextBox2.Text, TextBox1.Text, TextBox6.Text, TextBox3.Text, DropDownList1.Text, TextBox5.Text, numConf)
        emailSender.enviarEmail("uheredero001@gmail.com", TextBox1.Text, "Pincha aqui para confirmar registro: http://hads-18.azurewebsites.net/Confirmar.aspx?email=" + TextBox1.Text + "&numconf=" + numConf.ToString, "4QYzSiq7")
    End Sub

End Class