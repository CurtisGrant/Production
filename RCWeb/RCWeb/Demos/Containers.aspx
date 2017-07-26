<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Containers.aspx.vb" Inherits="RCWeb.Containers" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
       <asp:CheckBox ID="CheckBox1" runat="server" AutoPostBack="True" Text="Show Panel" OnCheckedChanged="CheckBox1_CheckedChanged"/>
       <asp:Panel ID="Panel1" runat="server" Visible="False">
          Can you see me now?</asp:Panel>
    
    </div>
    </form>
</body>
</html>
