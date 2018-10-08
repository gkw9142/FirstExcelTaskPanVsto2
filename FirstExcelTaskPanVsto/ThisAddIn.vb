Public Class ThisAddIn
    Private myUserControl1 As Calendar
    Private WithEvents myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myUserControl1 = New Calendar
        myCustomTaskPane = Me.CustomTaskPanes.Add(myUserControl1, "日历")

    End Sub

    Private Sub myCustomTaskPane_VisibleChanged(ByVal sender As Object,
                                                ByVal e As System.EventArgs) _
                Handles myCustomTaskPane.VisibleChanged

        Globals.Ribbons.ManageTaskPaneRibbon.ToggleButton1.Checked = myCustomTaskPane.Visible
    End Sub

    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return myCustomTaskPane
        End Get
    End Property

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
