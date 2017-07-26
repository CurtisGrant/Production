<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ListControls.aspx.vb" Inherits="RCWeb.ListControls" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
       <asp:DropDownList ID="DropDownList1" runat="server">
          <asp:ListItem>C#</asp:ListItem>
          <asp:ListItem>VB</asp:ListItem>
          <asp:ListItem>CSS</asp:ListItem>
       </asp:DropDownList>
       <asp:CheckBoxList ID="CheckBoxList1" runat="server">
          <asp:ListItem>C#</asp:ListItem>
          <asp:ListItem>VB</asp:ListItem>
          <asp:ListItem>CSS</asp:ListItem>
       </asp:CheckBoxList>
       <asp:Button ID="Button1" runat="server" Text="Button" Width="60px" />
       <br />
       <br />
       <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    
    </div>
    </form>
</body>
</html>
