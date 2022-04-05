Imports System.IO
Public Class _02

    Dim m As MsgBoxResult
    Dim t As Integer

    Private Sub MonthCalendar1_DateSelected(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        t = MonthCalendar1.SelectionRange.Start.Month.ToString & MonthCalendar1.SelectionRange.Start.Day.ToString
        Try
            If File.Exists(t & ".txt") = True Then
                MonthCalendar1.Enabled = False
                MonthCalendar1.Hide()
                TextBox1.Enabled = True
                TextBox1.Show()
                Button1.Enabled = True
                Button1.Show()
                Button2.Enabled = True
                Button2.Show()
                TextBox1.Text = File.ReadAllText(t & ".txt")
            Else
                m = MsgBox("would you like to enter appointments for this date?", MsgBoxStyle.YesNo)
                If m = MsgBoxResult.Yes Then
                    MonthCalendar1.Enabled = False
                    MonthCalendar1.Hide()
                    TextBox1.Enabled = True
                    TextBox1.Show()
                    TextBox1.Text = ""
                    Button1.Enabled = True
                    Button1.Show()
                    Button2.Enabled = True
                    Button2.Show()
              

                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        TextBox1.Enabled = False
        TextBox1.Hide()
        Button1.Enabled = False
        Button1.Hide()
        Button2.Enabled = False
        Button2.Hide()
        MonthCalendar1.Enabled = True
        MonthCalendar1.Show()

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Then
                If File.Exists(t & ".txt") = True Then
                    File.Delete(t & ".txt")
                End If
            End If
            If TextBox1.TextLength > 0 Then
                File.WriteAllText(t & ".txt", TextBox1.Text)


            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub _02_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim m1 As MsgBoxResult
        t = MonthCalendar1.SelectionRange.Start.Month.ToString & MonthCalendar1.SelectionRange.Start.Month.ToString
        If Date.Today = MonthCalendar1.TodayDate And File.Exists(t & ".txt") = True Then
            m1 = MsgBox("You have appointments set for today Would you like to view Them?", MsgBoxStyle.YesNo)
            If m1 - MsgBoxResult.Yes Then
                MonthCalendar1.Enabled = False
                MonthCalendar1.Hide()
                TextBox1.Enabled = True
                TextBox1.Show()
                Button1.Enabled = True
                Button1.Show()
                Button2.Enabled = True
                Button2.Show()
                TextBox1.Text = File.ReadAllText(t & ".txt")
            End If


        End If

    End Sub
End Class