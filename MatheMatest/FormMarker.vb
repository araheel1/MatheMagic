Public Class FormMarker
    Private MarkQuestions As New List(Of Integer)
    Private Count As Integer = -1
    Private MaxCount As Integer = -1
    Private Attempts As Integer

    Private Sub FormMarker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        MarkQuestions.Clear()
        MarkQuestions = Questions
        Count += 1
        MaxCount = MarkQuestions.Count - 1
        GetQInfo(Questions(Count))
        If Count = 0 Then
            Button6.Hide()
        Else
            Button6.Show()
        End If
        Colors()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label2.Text = TimeOfDay.ToString("h:mm tt")
    End Sub

    Private Sub Colors()
        If (ModuleID = 1) Or (ModuleID = 2) Then
            PictureBox10.BackColor = Color.PaleGreen
            Label4.BackColor = Color.PaleGreen
            Label4.ForeColor = Color.Black
            Label5.BackColor = Color.PaleGreen
            Label5.ForeColor = Color.Black
            Label6.BackColor = Color.PaleGreen
            Label6.ForeColor = Color.Black
            Label7.BackColor = Color.PaleGreen
            Label7.ForeColor = Color.Black
            Label8.BackColor = Color.PaleGreen
            Label8.ForeColor = Color.Black
            Label9.BackColor = Color.PaleGreen
            Label9.ForeColor = Color.Black
            Me.BackColor = Color.ForestGreen
            Button1.BackColor = Color.ForestGreen
            Button1.ForeColor = Color.White
            Button2.BackColor = Color.ForestGreen
            Button2.ForeColor = Color.White
            Button6.BackColor = Color.ForestGreen
            Button6.ForeColor = Color.White
            Label3.BackColor = Color.ForestGreen
            Label3.ForeColor = Color.White
            PictureBox8.Image = My.Resources._154910634887195497
            PictureBox8.BackColor = Color.PaleGreen
            Label10.BackColor = Color.ForestGreen
            Label10.ForeColor = Color.White
            PictureBox4.BackColor = Color.ForestGreen
            PictureBox5.BackColor = Color.ForestGreen
            Label1.BackColor = Color.ForestGreen
            Label2.BackColor = Color.ForestGreen
            PictureBox3.BackColor = Color.ForestGreen


        ElseIf ModuleID = 3 Then
            PictureBox10.BackColor = Color.CadetBlue
            Label4.BackColor = Color.CadetBlue
            Label4.ForeColor = Color.White
            Label5.BackColor = Color.CadetBlue
            Label5.ForeColor = Color.White
            Label6.BackColor = Color.CadetBlue
            Label6.ForeColor = Color.White
            Label7.BackColor = Color.CadetBlue
            Label7.ForeColor = Color.White
            Label8.BackColor = Color.CadetBlue
            Label8.ForeColor = Color.White
            Label9.BackColor = Color.CadetBlue
            Label9.ForeColor = Color.White
            Me.BackColor = Color.DarkCyan
            Button1.BackColor = Color.DarkCyan
            Button1.ForeColor = Color.White
            Button2.BackColor = Color.DarkCyan
            Button2.ForeColor = Color.White
            Button6.BackColor = Color.DarkCyan
            Button6.ForeColor = Color.White
            Label3.BackColor = Color.DarkCyan
            Label3.ForeColor = Color.White
            PictureBox8.Image = My.Resources.fp1boat
            PictureBox8.BackColor = Color.DarkCyan
            Label10.BackColor = Color.DarkCyan
            Label10.ForeColor = Color.White
            PictureBox4.BackColor = Color.DarkCyan
            PictureBox5.BackColor = Color.DarkCyan
            Label1.BackColor = Color.DarkCyan
            Label2.BackColor = Color.DarkCyan
            PictureBox3.BackColor = Color.DarkCyan


        ElseIf ModuleID = 4 Then
            PictureBox10.BackColor = Color.Moccasin
            Label4.BackColor = Color.Moccasin
            Label4.ForeColor = Color.Black
            Label5.BackColor = Color.Moccasin
            Label5.ForeColor = Color.Black
            Label6.BackColor = Color.Moccasin
            Label6.ForeColor = Color.Black
            Label7.BackColor = Color.Moccasin
            Label7.ForeColor = Color.Black
            Label8.BackColor = Color.Moccasin
            Label8.ForeColor = Color.Black
            Label9.BackColor = Color.Moccasin
            Label9.ForeColor = Color.Black
            Me.BackColor = Color.DarkOrange
            Button1.BackColor = Color.DarkOrange
            Button1.ForeColor = Color.White
            Button2.BackColor = Color.DarkOrange
            Button2.ForeColor = Color.White
            Button6.BackColor = Color.DarkOrange
            Button6.ForeColor = Color.White
            Label3.BackColor = Color.DarkOrange
            Label3.ForeColor = Color.White
            PictureBox8.Image = My.Resources.fm1arc
            PictureBox8.BackColor = Color.DarkOrange
            Label10.BackColor = Color.DarkOrange
            Label10.ForeColor = Color.White
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Count = MaxCount) Then
            MsgBox("End of question mark scheme.")
            Button1.Hide()
        Else
            Count += 1
            GetQInfo(Questions(Count))
        End If
        If Button6.Visible = False And Count > 0 Then
            Button6.Show()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Count = Count - 1
        GetQInfo(Questions(Count))
        If Count = 0 Then
            Button6.Hide()
        Else
            Button6.Show()
        End If
        Button1.Show()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Me.Close()
    End Sub

    Private Sub GetQInfo(QuestionIDx As Integer)
        Dim TempID As Integer
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From QuestionTopic, Questions Where (QuestionTopic.QuestionID = Questions.QuestionID) and (Questions.QuestionID = @QID)"
                .Parameters.AddWithValue("@QID", QuestionIDx)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                If rsp.Read Then
                    Label7.Text = rsp("ExamDate")
                    Label9.Text = rsp("QModule")
                    Label8.Text = rsp("Attempts")
                    Attempts = Label8.Text
                    TempID = rsp("TopicID")
                End If
                Label10.Text = "Question: " & Count + 1
                rsp.Close()
            End With
            cn.Close()
        End If
        QuestionID = QuestionIDx
        GetSubTopic(TempID)
        getMSImage(PictureBox1)
    End Sub

    Private Sub GetSubTopic(TopikID As Integer)
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Topic Where TopicID = @TID"
                .Parameters.AddWithValue("@TID", TopikID)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                If rsp.Read Then
                    Label1.Text = rsp("Topic") & ": " & rsp("SubTopic")
                End If
                rsp.Close()
            End With
            cn.Close()
        End If
    End Sub

    Private Sub getMSImage(pbx As PictureBox)
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd

                .Connection = cn
                .CommandText = "SELECT MSImage FROM Questions WHERE QuestionID = @ID"
                .Parameters.AddWithValue("@ID", QuestionID)

                Dim pictureData As Byte() = DirectCast(.ExecuteScalar(), Byte())

                'Create a stream in memory containing the bytes that comprise the image.
                Using stream As New IO.MemoryStream(pictureData)
                    'Read the stream and create an Image object from the data.
                    Dim pic As Image = Nothing
                    pic = Image.FromStream(stream)
                    pbx.Image = pic
                End Using

            End With

            cn.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PubAtt = Attempts
        FormFeedback.Show()
    End Sub

End Class