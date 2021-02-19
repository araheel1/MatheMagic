Public Class Form1
    Dim StartTime As DateTime 'old current time
    Dim PauseTime As DateTime 'keep track of the time when the timer is paused
    Dim TotalTimePaused As TimeSpan 'The total amount of time paused
    Dim ElapsedTime As TimeSpan

    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click
        ButtonStart.Enabled = False
        ButtonPause.Enabled = True
        StartTime = DateTime.Now()
        Timer1.Start()
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Subtract the current "Now" time from the start time
        'and then subtract the total time paused
        ElapsedTime = DateAndTime.Now.Subtract(StartTime).Subtract(TotalTimePaused)
        Label1.Text = ElapsedTime.Hours.ToString.PadLeft(2, "0"c) + ":" + ElapsedTime.Minutes.ToString.PadLeft(2, "0"c) + ":" + ElapsedTime.Seconds.ToString.PadLeft(2, "0"c) + "." + ElapsedTime.Milliseconds.ToString.PadLeft(3, "0"c)
        'And if you wanted the total elapsed time from time of start then
        'subtract the current "Now" time from the start time only.
    End Sub
    Private Sub ButtonPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPause.Click
        If Timer1.Enabled = False Then
            ButtonPause.Text = "Pause"
            'Add to the total time paused; the current time minus the 
            'time of pause.  this will be the total amount of time
            'paused so you can subtract it from the total amount of
            'time since the start button was pressed
            TotalTimePaused = TotalTimePaused.Add(DateAndTime.Now.Subtract(PauseTime))
            Timer1.Enabled = True
        Else
            ButtonPause.Text = "Resume"
            Timer1.Enabled = False
            'Set the pause time so it can be
            'calculated to the total time paused.
            'when the timer is resumed
            PauseTime = DateAndTime.Now
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ButtonPause.Text = "Resume" Then
            ButtonPause.PerformClick()
        End If
        StartTime += ElapsedTime
    End Sub
End Class