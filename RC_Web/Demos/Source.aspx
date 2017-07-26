<%@ Page Title="Source" Language="VB" MasterPageFile="~/MasterPages/Frontend.master" AutoEventWireup="false" 
    CodeFile="Source.aspx.vb" Inherits="Demos_Source" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cpMainContent" Runat="Server">
   
<asp:ListBox ID="ListBox1" runat="server" Width="150px" Height="166px"></asp:ListBox>
<br />
<hr />
<asp:TextBox ID="txtValue" runat="server" />
<asp:Button ID="btnAdd" Text="Add" runat="server" OnClientClick="return AddValues()" />
<script type="text/javascript">
    function AddValues() {
        var txtValue = document.getElementById("<%=txtValue.ClientID %>");
        var listBox = document.getElementById("<%= ListBox1.ClientID%>");
        var option = document.createElement("OPTION");
        option.innerHTML = txtValue.value;
        option.value = txtValue.value;
        if (option.value != "") {
        listBox.appendChild(option);}
    txtValue.value = "";
    return false;
}
</script>
 </asp:Content>
