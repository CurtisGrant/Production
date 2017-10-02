<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="VendorInventoryAnalysisReport.aspx.vb" Inherits="RCWeb.VendorInventoryAnalysisReport" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="height: 134px">
    
        <asp:Label ID="Label1" runat="server" Font-Size="Smaller" Text="Location" Width="50px"></asp:Label>
        <asp:DropDownList ID="ddStore" runat="server" Width="120px">
        </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label4" runat="server" Font-Size="Smaller" Text="Season" Width="60px"></asp:Label>
        <asp:DropDownList ID="ddSeason" runat="server" Width="120px">
        </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label7" runat="server" Font-Size="Smaller" Text="Report Date"></asp:Label>
        :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblReportDate" runat="server" Font-Size="Smaller" Text="Report Date"></asp:Label>
        <br />
        <asp:Label ID="Label2" runat="server" Font-Size="Smaller" Text="Dept" Width="50px"></asp:Label>
        <asp:DropDownList ID="ddDept" runat="server" AutoPostBack="True" Width="120px">
        </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label5" runat="server" Font-Size="Smaller" Text="Class" Width="60px"></asp:Label>
        <asp:DropDownList ID="ddClass" runat="server" Width="120px">
        </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblToday" runat="server" Font-Size="Smaller" Text="Today"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblDate4" runat="server" Font-Size="Smaller" Text="Next4"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblDate8" runat="server" Font-Size="Smaller" Text="Next8"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblDate12" runat="server" Font-Size="Smaller" Text="Next12"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblDate16" runat="server" Font-Size="Smaller" Text="Next16"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label8" runat="server" Font-Size="Smaller" Text="LY Sales:"></asp:Label>
&nbsp;<asp:Label ID="lblLYSales" runat="server" Font-Size="Smaller" Text="LYSales"></asp:Label>
        <br />
        <asp:Label ID="Label3" runat="server" Font-Size="Smaller" Text="Buyer" Width="50px"></asp:Label>
        <asp:DropDownList ID="ddBuyer" runat="server" Width="120px">
        </asp:DropDownList>
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label6" runat="server" Font-Size="Smaller" Text="Min OH$" Width="60px"></asp:Label>
        <asp:TextBox ID="txtMinOH" runat="server" Width="108px"></asp:TextBox>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblNow" runat="server" Font-Bold="True" Font-Size="Smaller" Text="Today"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblNex4" runat="server" Font-Bold="True" Font-Size="Smaller" Text="Next4"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblnNex8" runat="server" Font-Bold="True" Font-Size="Smaller" Text="Next8"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblNex12" runat="server" Font-Bold="True" Font-Size="Smaller" Text="Next12"></asp:Label>
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="lblNex16" runat="server" Font-Bold="True" Font-Size="Smaller" Text="Next16"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label9" runat="server" Font-Size="Smaller" Text="TY Sales:"></asp:Label>
&nbsp;<asp:Label ID="lblTYSales" runat="server" Font-Size="Smaller" Text="TYSales"></asp:Label>
        <br />
        <asp:Button ID="btnRunReport" runat="server" Text="Run Report" />
        <br />
        <br />
        <asp:GridView ID="dgv1" runat="server">
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
