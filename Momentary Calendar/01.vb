Public Class Form1


    Private Sub DateTimePicker2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        ' check date picker
        If DateTimePicker2.Value < DateTimePicker1.Value Then
            MsgBox("DateTimePicker2 smallest. try again.")
        End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        If DateTimePicker2.Value > DateTimePicker1.Value Then
            MsgBox("DateTimePicker2 Bigger. try again.")
        End If
    End Sub
End Class
