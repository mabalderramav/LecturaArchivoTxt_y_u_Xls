<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="GeoPlus.CargaArchivoWeb.index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FileUpload ID="fuPrueba" runat="server" />
&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnProcesar" runat="server" OnClick="btnProcesar_Click" Text="Procesar" />
&nbsp;&nbsp;
        <asp:Label ID="lblMensaje" runat="server" Text="esto es un caso de pureba :)"></asp:Label>
    
    </div>
    </form>
</body>
</html>
