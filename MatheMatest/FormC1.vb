Public Class FormC1

    Private Sub DisplayRQuestions(TopikID As Integer)
        ComboBox1.Items.Clear()
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From QuestionTopic Where (TopicID = @TID) and (FeedbackDate <= @Date)"
                .Parameters.AddWithValue("@TID", TopikID)
                .Parameters.AddWithValue("@Date", Today.Date)
                TopicID = TopikID
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                While rsp.Read
                    Dim ShowValue As String = "Question" & " " & rsp("QuestionID")
                    Dim CustomerValue As New ListBoxData(ShowValue, rsp("QuestionID"))
                    ComboBox1.Items.Add(CustomerValue)
                End While
                rsp.Close()
            End With
            cn.Close()
        End If
        If ComboBox1.Items.Count = 0 Then
            PictureBox4.Visible = False
            MsgBox("No questions currently available for the selected topic, please select an alternative topic.")
        Else
            PictureBox4.Visible = True
        End If
    End Sub

    Private Sub DisplayQuestion(QuestionxID As Integer)
        QuestionID = QuestionxID
        FormQandA.Show()
    End Sub



    Private Sub FormC1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Look for a match and display it in the combobox
        If DbConnect() Then
            LstSelectCust.Items.Clear()
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Topic Where ModuleID = @MID"
                .Parameters.AddWithValue("@MID", 1)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                While rsp.Read
                    Dim ShowValue As String = rsp("Topic") & ": " & rsp("SubTopic")
                    Dim CustomerValue As New ListBoxData(ShowValue, rsp("TopicID"))
                    LstSelectCust.Items.Add(CustomerValue)
                End While
                rsp.Close()
            End With
            cn.Close()
        End If
    End Sub

    Private Sub LstSelectCust_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstSelectCust.SelectedIndexChanged
        ComboBox1.Text = ""
        If LstSelectCust.SelectedItem IsNot Nothing Then
            Dim SelectedCust As ListBoxData = LstSelectCust.SelectedItem
            DisplayRQuestions(SelectedCust.Identifier)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            Dim SelectedCust As ListBoxData = ComboBox1.SelectedItem
            DisplayQuestion(SelectedCust.Identifier)
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim rnd As New Random
        Dim randomIndex As Integer = rnd.Next(0, LstSelectCust.Items.Count - 1)
        LstSelectCust.SelectedItem = (LstSelectCust.Items(randomIndex))
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Dim rnd As New Random
        Dim randomIndex As Integer = rnd.Next(0, ComboBox1.Items.Count - 1)
        ComboBox1.SelectedItem = (ComboBox1.Items(randomIndex))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ModuleID = 1
        FormPaperSet.Show()
    End Sub
End Class