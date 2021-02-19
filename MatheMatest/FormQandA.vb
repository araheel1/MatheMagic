Public Class FormQandA
    Private Attempts As Integer
    Dim MyName As String = FormSelectBook.Name
    Dim StartTime As DateTime
    Dim Elapsed As TimeSpan
    Dim penselect As Boolean = False
    Dim RanNum As New Random
    Public Picture As Image
    Dim TimeOn As Boolean = False
    Dim TempQID As Integer = -1

    Private Sub FormQandA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
        StartTime = Now
        GetTopic()
        PictureBox2.Hide()
        PictureBox6.Image = My.Resources.pencil2
        PictureBox11.Image = My.Resources.timer2
        Colors()
        Button3.Hide()
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
            Button3.BackColor = Color.ForestGreen
            Button3.ForeColor = Color.White
            Label3.BackColor = Color.ForestGreen
            Label3.ForeColor = Color.White
            Label11.BackColor = Color.ForestGreen
            Label11.ForeColor = Color.White
            PictureBox8.Image = My.Resources._154910634887195497
            PictureBox8.BackColor = Color.PaleGreen
            Label10.BackColor = Color.ForestGreen
            Label10.ForeColor = Color.White
            PictureBox4.BackColor = Color.ForestGreen
            PictureBox5.BackColor = Color.ForestGreen
            PictureBox7.BackColor = Color.ForestGreen
            PictureBox11.BackColor = Color.ForestGreen
            PictureBox6.BackColor = Color.ForestGreen
            PictureBox12.BackColor = Color.ForestGreen
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
            Button3.BackColor = Color.DarkCyan
            Button3.ForeColor = Color.White
            Label3.BackColor = Color.DarkCyan
            Label3.ForeColor = Color.White
            Label11.BackColor = Color.DarkCyan
            Label11.ForeColor = Color.White
            PictureBox8.Image = My.Resources.fp1boat
            PictureBox8.BackColor = Color.DarkCyan
            Label10.BackColor = Color.DarkCyan
            Label10.ForeColor = Color.White
            PictureBox4.BackColor = Color.DarkCyan
            PictureBox5.BackColor = Color.DarkCyan
            PictureBox7.BackColor = Color.DarkCyan
            PictureBox11.BackColor = Color.DarkCyan
            PictureBox6.BackColor = Color.DarkCyan
            PictureBox12.BackColor = Color.DarkCyan
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
            Button3.BackColor = Color.DarkOrange
            Button3.ForeColor = Color.White
            Label3.BackColor = Color.DarkOrange
            Label3.ForeColor = Color.White
            PictureBox8.Image = My.Resources.fm1arc
            PictureBox8.BackColor = Color.DarkOrange
            Label10.BackColor = Color.DarkOrange
            Label10.ForeColor = Color.White
            Label11.BackColor = Color.DarkOrange
            Label11.ForeColor = Color.White
        End If
    End Sub

    Private Sub GetTopic()
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Topic Where TopicID = @TID"
                .Parameters.AddWithValue("@TID", TopicID)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                If rsp.Read Then
                    Label1.Text = rsp("SubTopic")
                End If
                rsp.Close()
            End With
            cn.Close()
        End If
        GetQInfo()
    End Sub

    Private Sub GetQInfo()
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Questions Where QuestionID = @QID"
                .Parameters.AddWithValue("@QID", QuestionID)
                Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                If rsp.Read Then
                    Label7.Text = rsp("ExamDate")
                    Label9.Text = rsp("QModule")
                    Label8.Text = rsp("Attempts")
                    Attempts = Label8.Text
                End If
                rsp.Close()
            End With
            cn.Close()
        End If
        getQImage(PictureBox1)

        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd

                .Connection = cn
                .CommandText = "SELECT MSImage FROM Questions WHERE QuestionID = @ID"
                .Parameters.AddWithValue("@ID", QuestionID)

                Dim pictureData As Byte() = DirectCast(.ExecuteScalar(), Byte())

                Dim picture As Image = Nothing

                'Create a stream in memory containing the bytes that comprise the image.
                Using stream As New IO.MemoryStream(pictureData)
                    'Read the stream and create an Image object from the data.
                    picture = Image.FromStream(stream)
                    PictureBox2.Image = picture
                End Using
            End With
            cn.Close()
        End If
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label2.Text = TimeOfDay.ToString("h:mm tt")
        If TimeOn = True Then
            Elapsed = Now - StartTime
            Label3.Text = "You're on the clock: " & Elapsed.Minutes & " minute(s) and " & Elapsed.Seconds & " second(s)!"
        End If
    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        If TimeOn = True Then
            TimeOn = False
            PictureBox11.Image = My.Resources.timer2
            Timer1.Enabled = False
        Else
            TimeOn = True
            PictureBox11.Image = My.Resources.timergreen
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        PictureBox2.Show()
        Button3.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PictureBox2.Hide()
        Button3.Hide()
    End Sub

    Private _Previous As System.Nullable(Of Point) = Nothing
    Private Sub pictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        _Previous = e.Location
        pictureBox1_MouseMove(sender, e)
    End Sub

    Private Sub pictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If penselect = True Then
            If _Previous IsNot Nothing Then
                If PictureBox1.Image Is Nothing Then
                    Dim bmp As New Bitmap(PictureBox1.Width, PictureBox1.Height)
                    Using g As Graphics = Graphics.FromImage(bmp)
                        g.Clear(Color.White)
                    End Using
                    PictureBox1.Image = bmp
                End If
                Using g As Graphics = Graphics.FromImage(PictureBox1.Image)
                    g.DrawLine(Pens.Black, _Previous.Value, e.Location)
                End Using
                PictureBox1.Invalidate()
                _Previous = e.Location
            End If
        End If
    End Sub

    Private Sub pictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        _Previous = Nothing
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        If penselect = False Then
            penselect = True
            PictureBox6.Image = My.Resources.pencilgreen
        Else
            penselect = False
            PictureBox6.Image = My.Resources.pencil2
        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label10.Text = "Your hair is looking" & vbNewLine & "exquitsite as always, " & MyName & "!"
            Case 2
                Label10.Text = "You ARE formidable." & vbNewLine & "Don't let anyone" & vbNewLine & "tell you otherwise " & MyName & "."
            Case 3
                Label10.Text = "It could be worse." & vbNewLine & "You could be doing" & vbNewLine & "Banter (Decision) Maths!"
            Case 4
                Label10.Text = "You can tear through these" & vbNewLine & "problems like it's" & vbNewLine & "no one's business, " & MyName & "!"
        End Select
    End Sub

    Private Sub Label10_MouseHover(sender As Object, e As EventArgs) Handles Label10.MouseHover
        Label10.ForeColor = Color.Silver
    End Sub

    Private Sub FormQandA_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseHover
        Label10.ForeColor = Color.White
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label10.Text = "Click me for" & vbNewLine & "words of encouragement!"
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        Picture = PictureBox1.Image
        FormQ.Show()
    End Sub

    Public Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        Using g As Graphics = Graphics.FromImage(PictureBox1.Image)
            g.Clear(Color.White)
        End Using
        PictureBox1.Invalidate()
        getQImage(PictureBox1)
        GetMSImage()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PubAtt = Attempts
        FormFeedback.Show()
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Dim k As New Graphing.FormGraphs
        k.Show()
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        Picture = PictureBox2.Image
        FormM.Show()
    End Sub

    Private Sub GetMSImage()
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd

                .Connection = cn
                .CommandText = "SELECT MSImage FROM Questions WHERE QuestionID = @ID"
                .Parameters.AddWithValue("@ID", QuestionID)

                Dim pictureData As Byte() = DirectCast(.ExecuteScalar(), Byte())

                Dim picture As Image = Nothing

                'Create a stream in memory containing the bytes that comprise the image.
                Using stream As New IO.MemoryStream(pictureData)
                    'Read the stream and create an Image object from the data.
                    picture = Image.FromStream(stream)
                    PictureBox2.Image = picture
                End Using
            End With
            cn.Close()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TimeOn = False
        StartTime = Now
        Label3.Text = "Timer reset. Hit the timer when you're ready!"
    End Sub
End Class