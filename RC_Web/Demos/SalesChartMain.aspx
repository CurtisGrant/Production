<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SalesChartMain.aspx.vb" Inherits="Demos_SalesChartMain" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="Panel1" runat="server">
            Select Date<br />
        <asp:TextBox ID="txtDate" runat="server" Height="22px" Width="129px"></asp:TextBox>
            <ajaxToolkit:CalendarExtender ID="txtDate_CalendarExtender" runat="server" BehaviorID="txtDate_CalendarExtender" TargetControlID="txtDate" />
        <asp:ImageButton ID="ImageButton1" imageurl="~/Images/calendar.jpg" runat="server" Height="30px" Width="30px" />
        
            </asp:Panel>
    <p />

    </div>
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
    </form>
</body>
</html>
