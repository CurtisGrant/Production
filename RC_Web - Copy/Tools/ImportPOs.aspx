<%@ Page Title="Import POs" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="ImportPOs.aspx.vb" Inherits="Tools_ImportPOs" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .auto-style1 {
            width: 100%;
        }
    .auto-style2 {
        height: 205px;
    }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
   
    <table class="auto-style1">
        <tr>
            <td>
        <asp:Label ID="Label2" runat="server" Font-Size="Small" Text="Select Purchase Request"></asp:Label>
            </td>
            <td>
    <asp:FileUpload ID="FileUpload1" runat="server" Font-Size="Small" Width="200px" />
            </td>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td>
        <asp:DropDownList ID="ddPreq" runat="server" Font-Size="Small" Height="19px" Width="323px">
        </asp:DropDownList>
            </td>
            <td>
    <asp:Label ID="Label1" runat="server" Font-Size="Small"></asp:Label>
            </td>
            <td>
    <asp:Button ID="Button1" runat="server" Text="Load File" Font-Size="Small" />
            </td>
        </tr>
    </table>

    <table class="auto-style1">
        <tr>
            <td class="auto-style2">
                <asp:ListBox ID="RequiredFields" runat="server" Height="174px">
                    <asp:ListItem>Item</asp:ListItem>
                    <asp:ListItem>Color</asp:ListItem>
                    <asp:ListItem>Size</asp:ListItem>
                    <asp:ListItem>Size2</asp:ListItem>
                    <asp:ListItem>Description</asp:ListItem>
                    <asp:ListItem>Cost</asp:ListItem>
                    <asp:ListItem>Retail</asp:ListItem>
                    <asp:ListItem>Department</asp:ListItem>
                    <asp:ListItem>Class</asp:ListItem>
                    <asp:ListItem>Unit of Measure</asp:ListItem>
                </asp:ListBox>
                <br />
                <asp:Label ID="Label3" runat="server" Text="Required Fields"></asp:Label>
            </td>

            <td class="auto-style2">

                <asp:ListBox ID="YourFields" runat="server" AutoPostBack="True" Height="168px" Width="114px"></asp:ListBox>
                <br />
                <asp:Label ID="Label4" runat="server" Text="Your Fields"></asp:Label>
            </td>
            <td class="auto-style2">
                <asp:ListBox ID="BoundFields" runat="server" Height="171px" Width="115px"></asp:ListBox>
                <br />
                <asp:Label ID="Label5" runat="server" Text="Bound Fields"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                &nbsp;</td>
        </tr>
    </table>
    <br />
    <p>
        &nbsp;</p>
    <asp:GridView ID="GridView1" runat="server" Font-Size="Small" SkinID="Professional" OnRowDataBound="gridview1_RowDataBound"
        OnRowCommand="gridview1_RowCommand" >
        <AlternatingRowStyle BackColor="#CCCCCC" Font-Size="Small" />
        
    </asp:GridView>
    <p>
        &nbsp;</p>
</asp:Content>

