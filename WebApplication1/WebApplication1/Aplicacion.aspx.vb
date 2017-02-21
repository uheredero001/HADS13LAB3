Public Class Aplicacion
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("user") Is Nothing) Then
            Response.Redirect("Inicio.aspx")
        End If
    End Sub

End Class