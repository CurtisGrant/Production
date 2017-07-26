
Partial Class Demos_Source:
    Inherits BasePage

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        ''Server.Transfer("Target.aspx?Test=SomeValue")
        If Not IsPostBack Then ListBox1.Items.Clear()
        ListBox1.Items.Add("A")
        ListBox1.Items.Add("B")
        ListBox1.Items.Add("C")
    End Sub

    Protected Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub


End Class
