Public Class _03

    Private Sub _03_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        tmrSync.Enabled = True


    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub tmrSync_Tick(sender As System.Object, e As System.EventArgs) Handles tmrSync.Tick
        Label4.Text = TimeOfDay
        Label3.Text = Date.Today


    End Sub
End Class