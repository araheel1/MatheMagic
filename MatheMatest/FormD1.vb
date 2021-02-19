Public Class FormD1

    Private Sub DisplayRQuestions(TopicID As Integer)
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From QuestionTopic Where TopicID = @TID"
                .Parameters.AddWithValue("@TID", TopicID)
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
    End Sub

    Private Sub DisplayQuestion(QuestionID As Integer)
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Questions Where QuestionID = @QID"
                .Parameters.AddWithValue("@QID", QuestionID)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()
                If rsp.Read Then
                    TextBox1.Text = rsp("ExamUnit")
                End If
                rsp.Close()
            End With
            cn.Close()
        End If
    End Sub



    Private Sub FormD1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Look for a match and display it in the combobox
        If DbConnect() Then
            LstSelectCust.Items.Clear()
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Topic Where ModuleID = @MID"
                .Parameters.AddWithValue("@MID", 5)
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
End Class