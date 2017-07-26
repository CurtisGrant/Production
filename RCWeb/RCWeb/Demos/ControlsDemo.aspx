<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ControlsDemo.aspx.vb" Inherits="RCWeb.ControlsDemo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
       Yo Name<asp:TextBox ID="YoName" runat="server" style="margin-left: 14px"></asp:TextBox>
       <asp:Button ID="Sumbit" runat="server" Text="Sumbit" />
       <br />
       <asp:Label ID="Result" runat="server"></asp:Label>
    </form>
</body>
</html>
