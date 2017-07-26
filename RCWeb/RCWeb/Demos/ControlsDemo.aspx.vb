Public Class ControlsDemo
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

   Protected Sub Sumbit_Click(sender As Object, e As EventArgs) Handles Sumbit.Click
      Result.Text = "Yo name is " & YoName.Text
   End Sub
End Class