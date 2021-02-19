Public Class FormFeedback
    Dim rate As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "UPDATE QuestionTopic SET FeedbackDate = @Date, FeedbackInfo = @Info WHERE QuestionID = @ID"
                .Parameters.AddWithValue("@Date", Today.Date.AddDays(Math.Round(rate / 2)))
                .Parameters.AddWithValue("@Info", rate)
                .Parameters.AddWithValue("@ID", QuestionID)
                .ExecuteNonQuery()
            End With
        End If
        cn.Close()

        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "UPDATE Questions SET Attempts = @Att WHERE QuestionID = @ID"
                .Parameters.AddWithValue("@Att", PubAtt + 1)
                .Parameters.AddWithValue("@ID", QuestionID)
                .ExecuteNonQuery()
            End With
        End If
        cn.Close()

        MsgBox("Thank you for adding feedback.")
        Me.Close()
    End Sub

    Private Sub GetRate()
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From QuestionTopic Where QuestionID = @QID"
                .Parameters.AddWithValue("@QID", QuestionID)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()
                If rsp.Read Then
                    rate = rsp("FeedbackInfo")
                End If
                rsp.Close()
            End With
            cn.Close()
        End If
        Label7.Text = rate & "/10"
    End Sub

    Private Sub Colors()
        If (ModuleID = 1) Or (ModuleID = 2) Then
            PictureBox4.BackColor = Color.ForestGreen
            PictureBox5.BackColor = Color.ForestGreen
            PictureBox5.Image = My.Resources._154910634887195497
            Label3.BackColor = Color.ForestGreen
            Label3.ForeColor = Color.White
            Label4.BackColor = Color.ForestGreen
            Label4.ForeColor = Color.White
            Label5.BackColor = Color.ForestGreen
            Label5.ForeColor = Color.White
            Label6.BackColor = Color.ForestGreen
            Label6.ForeColor = Color.White
            Label7.BackColor = Color.ForestGreen
            Label7.ForeColor = Color.White
            Button1.BackColor = Color.ForestGreen
            Button1.ForeColor = Color.White
            Label1.BackColor = Color.PaleGreen
            Label1.ForeColor = Color.Black
            Label2.BackColor = Color.PaleGreen
            Label2.ForeColor = Color.Black
            Me.BackColor = Color.PaleGreen
        ElseIf ModuleID = 3 Then
            PictureBox4.BackColor = Color.DarkCyan
            PictureBox5.BackColor = Color.DarkCyan
            PictureBox5.Image = My.Resources.fp1boat
            Label3.BackColor = Color.DarkCyan
            Label3.ForeColor = Color.White
            Label4.BackColor = Color.DarkCyan
            Label4.ForeColor = Color.White
            Label5.BackColor = Color.DarkCyan
            Label5.ForeColor = Color.White
            Label6.BackColor = Color.DarkCyan
            Label6.ForeColor = Color.White
            Label7.BackColor = Color.DarkCyan
            Label7.ForeColor = Color.White
            Button1.BackColor = Color.DarkCyan
            Button1.ForeColor = Color.White
            Label1.BackColor = Color.CadetBlue
            Label1.ForeColor = Color.White
            Label2.BackColor = Color.CadetBlue
            Label2.ForeColor = Color.White
            Me.BackColor = Color.CadetBlue
        End If
    End Sub

    Private Sub FormFeedback_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = PubAtt
        Label6.Text = "Question " & QuestionID
        GetRate()
        Colors()
        TrackBar1.Value = 5
        Button1.Hide()
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        TextBox1.Text = TrackBar1.Value
        rate = TextBox1.Text
        Button1.Show()
    End Sub
End Class