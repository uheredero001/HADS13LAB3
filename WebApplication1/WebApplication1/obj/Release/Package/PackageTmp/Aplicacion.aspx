<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Aplicacion.aspx.vb" Inherits="WebApplication1.Aplicacion" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="height: 222px; background-color: #99CCFF;">
    
        <asp:Label ID="Label1" runat="server" Text="Hola!!!!" Font-Size="X-Large"></asp:Label>
        <br />
        <asp:Image ID="Image1" runat="server" Height="192px" ImageUrl="~/Koala.jpg" Width="248px" />
    
    </div>
    </form>
</body>
</html>
