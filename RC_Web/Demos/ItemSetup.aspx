<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPages/ItemManagement.master" AutoEventWireup="false" CodeFile="ItemSetup.aspx.vb" Inherits="Demos_ItemSetup" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .auto-style1 {
            width: 100%;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
    <h2>Setup New Items</h2>

    <table class="auto-style1">
        <tr>
            <td>Item</td>
            <td>
                <asp:TextBox ID="txtItem" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Description</td>
            <td>
                <asp:TextBox ID="txtDesc" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Vendor</td>
            <td>
                <asp:TextBox ID="txtVendor" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Category</td>
            <td>
                <asp:TextBox ID="txtCateg" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>SubCategory</td>
            <td>
                <asp:TextBox ID="txtSubCateg" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Buyer</td>
            <td>
                <asp:TextBox ID="txtBuyer" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Unit of Measure</td>
            <td>
                <asp:TextBox ID="txtUOM" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Buy Unit</td>
            <td>
                <asp:TextBox ID="txtBuyUnit" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>Sell Unit</td>
            <td>
                <asp:TextBox ID="txtSellUnit" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td></td>
            <td>
                <asp:Button ID="btnSave" runat="server" Text="Button" />
            </td>
            
        </tr>
    </table>
        <p />

        <asp:ListView ID="lv1" runat="server">
        </asp:ListView>
</asp:Content>

