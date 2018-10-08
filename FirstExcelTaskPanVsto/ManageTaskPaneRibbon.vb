Imports Microsoft.Office.Tools.Ribbon

Public Class ManageTaskPaneRibbon

    Private Sub ManageTaskPaneRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ToggleButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton1.Click
        Globals.ThisAddIn.TaskPane.Visible =
            TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub
End Class
