<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Confirmar.aspx.vb" Inherits="WebApplication1.Confirmar" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" style="background-color: #99CCFF">
        <asp:Label ID="Label1" runat="server" Text="Registro confirmado!!" Font-Size="Larger"></asp:Label>
        <br />
        <asp:Image ID="Image1" runat="server" ImageUrl="~/okai.jpg" />
        <asp:Label ID="Label2" runat="server"></asp:Label>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Inicio.aspx">Inicio</asp:HyperLink>
    </form>
</body>
</html>
