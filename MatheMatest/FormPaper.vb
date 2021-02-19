Public Class FormPaper
    Public marks As Integer
    Private Count As Integer = -1
    Private MaxCount As Integer = -1
    Dim TeeQs As List(Of Integer)
    Dim RanNum As New Random
    Dim TimeOn As Boolean = True
    Dim MyName As String = FormSelectBook.Name
    Dim StartTime As DateTime
    Dim Elapsed As TimeSpan
    Dim penselect As Boolean = False
    Dim open As Boolean = False
    Dim complete As Boolean = False
    Dim tempMark As Integer
    Dim Tried As Boolean

    Private Sub GetTopics()
        ListBox2.Items.Clear()
        For x As Integer = 0 To ListBox1.Items.Count - 1
            If DbConnect() Then
                Dim SQLCmd As New OleDb.OleDbCommand
                With SQLCmd
                    .Connection = cn
                    .CommandText = "Select * From QuestionTopic Where TopicID = @TID and FeedbackDate <= @Date"
                    .Parameters.AddWithValue("@TID", ListBox1.Items.Item(x))
                    .Parameters.AddWithValue("@Date", Today.Date)
                    Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                    While rsp.Read
                        Dim countx As Integer = 0
                        Dim Temp As Integer = rsp("QuestionID")
                        For y As Integer = 0 To ListBox2.Items.Count - 1
                            If (Temp = ListBox2.Items.Item(y)) Then
                                countx += 1
                            End If
                        Next y
                        If countx > 0 Then
                            'Don't add the question
                        Else
                            ListBox2.Items.Add(Temp)
                        End If
                    End While
                    rsp.Close()
                End With
                cn.Close()
            End If
        Next x
        If ListBox2.Items.Count = 0 Then
            MsgBox("No questions currently available for the selected topics... Please reselect topics.")
            open = True
        End If
        If open = False Then
            GetMarks()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub GetMarks()
        ListBox4.Items.Clear()
        For x As Integer = 0 To ListBox2.Items.Count - 1
            If DbConnect() Then
                Dim SQLCmd As New OleDb.OleDbCommand
                With SQLCmd
                    .Connection = cn
                    .CommandText = "Select * From Questions Where QuestionID = @QID"
                    .Parameters.AddWithValue("@QID", ListBox2.Items.Item(x))
                    Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

                    While rsp.Read
                        Dim Temp As Integer = rsp("QMarkNum")
                        ListBox4.Items.Add(Temp)
                    End While
                    rsp.Close()
                End With
                cn.Close()
            End If
        Next x
        Tried = False
        GetQuestion()
    End Sub

    'Private Sub GetQuestion()
    '    If (Convert.ToInt32(Label8.Text) >= marks - 5) Then
    '        MsgBox("End of test - mark threshold reached.")
    '        Button3.Hide()
    '        FormMarker.Show()
    '        complete = True
    '    ElseIf ListBox2.Items.Count = 0 Then
    '        MsgBox("End of test - no more questions remaining for the selected topics.")
    '        Button3.Hide()
    '        FormMarker.Show()
    '        complete = True
    '    Else
    '        Dim myRandom As New Random
    '        Dim i As Integer = ListBox2.Items.Count
    '        Dim n As Integer = ListBox2.Items.Item(myRandom.Next(i))
    '        ListBox3.Items.Add(n)
    '        Questions.Add(n)
    '        ListBox2.Items.Remove(n)
    '        Count += 1
    '        If (Count > MaxCount) Then
    '            MaxCount = Count
    '        End If
    '        GetQInfo(Questions(Count))
    '        If Count = 0 Then
    '            Button6.Hide()
    '        Else
    '            Button6.Show()
    '        End If
    '    End If
    'End Sub

    Private Sub GetQuestion()
        If ListBox4.Items.Count = 0 And Tried = True Then
            MsgBox("End of test - mark threshold reached.")
            Button3.Hide()
            FormMarker.Show()
            complete = True
        ElseIf ListBox4.Items.Count = 0 And Tried = False Then
            MsgBox("End of test - no more questions currently avaiable for the selected topics.")
            Button3.Hide()
            FormMarker.Show()
            complete = True
        Else
            Tried = True
            Dim myRandom As New Random
            Dim i As Integer = ListBox4.Items.Count
            Dim n As Integer = ListBox4.Items.Item(myRandom.Next(i))
            MsgBox(n)
            If CInt(Label8.Text) + n <= marks Then
                MsgBox("true")
                Dim TQID As Integer = ListBox2.Items.Item(ListBox4.Items.IndexOf(n))
                MsgBox(TQID)
                ListBox4.Items.Remove(n)
                ListBox3.Items.Add(TQID)
                Questions.Add(TQID)
                ListBox2.Items.Remove(TQID)
                Count += 1
                If (Count > MaxCount) Then
                    MaxCount = Count
                End If
                GetQInfo(Questions(Count))
                If Count = 0 Then
                    Button6.Hide()
                Else
                    Button6.Show()
                End If
            Else
                Dim TQIDx As Integer = ListBox2.Items.Item(ListBox4.Items.IndexOf(n))
                ListBox2.Items.Remove(TQIDx)
                ListBox4.Items.Remove(n)
                GetQuestion()
            End If
        End If
    End Sub

    Private Sub FormPaper_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Questions.Clear()
        ListBox1.Items.Clear()
        ListBox1.DataSource = itemList
        GetTopics()
        Timer1.Start()
        StartTime = Now
        PictureBox6.Image = My.Resources.pencil2
        Colors()
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
            Button3.BackColor = Color.PaleGreen
            Button3.ForeColor = Color.Black
            Me.BackColor = Color.ForestGreen
            Button1.BackColor = Color.ForestGreen
            Button1.ForeColor = Color.White
            Button2.BackColor = Color.ForestGreen
            Button2.ForeColor = Color.White
            Button3.BackColor = Color.ForestGreen
            Button3.ForeColor = Color.White
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
            PictureBox7.BackColor = Color.ForestGreen
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
            Button3.BackColor = Color.CadetBlue
            Button3.ForeColor = Color.White
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
            PictureBox7.BackColor = Color.DarkCyan
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
            Button3.BackColor = Color.Moccasin
            Button3.ForeColor = Color.Black
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

    'Private Sub GetTopic()
    '    Dim myRandom As New Random
    '    Dim i As Integer = ListBox1.Items.Count
    '    Dim SelectedVal As Integer = ListBox1.Items.Item(myRandom.Next(i))
    '    GetRQuestions(SelectedVal)
    'End Sub

    'Private Sub GetRQuestions(TopikID As Integer)
    '    ListBox2.Items.Clear()
    '    If DbConnect() Then
    '        Dim SQLCmd As New OleDb.OleDbCommand
    '        With SQLCmd
    '            .Connection = cn
    '            .CommandText = "Select * From QuestionTopic Where TopicID = @TID"
    '            .Parameters.AddWithValue("@TID", TopikID)
    '            TopicID = TopikID
    '            Dim rsp As OleDb.OleDbDataReader = .ExecuteReader()

    '            While rsp.Read
    '                Dim ShowValue As String = rsp("QuestionID")
    '                Dim CustomerValue As New ListBoxData(ShowValue, rsp("TopicID"))
    '                ListBox2.Items.Add(CustomerValue)
    '            End While
    '            rsp.Close()
    '        End With
    '        cn.Close()
    '    End If
    '    GetQuestion(TopikID)
    'End Sub

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

    'Private Sub GetQuestion(TopikIDa As Integer)
    '    Dim myRandom As New Random
    '    Dim i As Integer = ListBox2.Items.Count
    '    Dim count3 As Integer = 0
    '    Dim SelectedVal As ListBoxData = ListBox2.Items.Item(myRandom.Next(i))
    '    MsgBox(SelectedVal.Data)
    '    MsgBox(SelectedVal.Identifier)
    '    For x As Integer = 0 To ListBox3.Items.Count - 1
    '        Dim k As ListBoxData = ListBox3.Items.Item(x)
    '        MsgBox(k.Data)
    '        MsgBox(k.Identifier)
    '        If ((k.Data) = (SelectedVal.Data)) And (k.Identifier = SelectedVal.Identifier) Then
    '            count3 += 1
    '        End If
    '    Next x
    '    If ListBox3.Items.Count - 1 = ListBox1.Items.Count - 1 Then
    '        'MsgBox("wyd fam")
    '        FormMarker.Show()
    '    ElseIf count3 > 0 Then
    '        GetTopic()
    '    ElseIf count3 = 0 Then
    '        Dim ShowValue As String = SelectedVal.Data
    '        Dim QValue As New ListBoxData(ShowValue, TopikIDa)
    '        ListBox3.Items.Add(QValue)
    '        Questions.Add(SelectedVal.Data)
    '        Count += 1
    '        If (Count > MaxCount) Then
    '            MaxCount = Count
    '        End If
    '        GetQInfo(Questions(Count))
    '        If Count = 0 Then
    '            Button6.Hide()
    '        Else
    '            Button6.Show()
    '        End If
    '    End If
    'End Sub

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
                    If Count = MaxCount And complete = False Then
                        tempMark = rsp("QMarkNum")
                        MsgBox("TMark: " & tempMark)
                        Label8.Text = Convert.ToInt32(Label8.Text) + tempMark
                    End If
                    TempID = rsp("TopicID")
                End If
                Label10.Text = "Question: " & Count + 1
                rsp.Close()
            End With
            cn.Close()
        End If
        QuestionID = QuestionIDx
        GetSubTopic(TempID)
        getQImage(PictureBox1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Count = MaxCount) Then
            Button3.Show()
            Tried = False
            GetQuestion()
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
        Button3.Hide()
        If Count = 0 Then
            Button6.Hide()
        Else
            Button6.Show()
        End If
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label11.Text = "Your hair is looking" & vbNewLine & "exquitsite as always, " & MyName & "!"
            Case 2
                Label11.Text = "You ARE formidable." & vbNewLine & "Don't let anyone" & vbNewLine & "tell you otherwise " & MyName & "."
            Case 3
                Label11.Text = "It could be worse." & vbNewLine & "You could be doing" & vbNewLine & "Banter (Decision) Maths!"
            Case 4
                Label11.Text = "You can tear through these" & vbNewLine & "problems like it's" & vbNewLine & "no one's business, " & MyName & "!"
        End Select
    End Sub

    Private Sub Label11_MouseHover(sender As Object, e As EventArgs) Handles Label11.MouseHover
        Label11.ForeColor = Color.Silver
    End Sub

    Private Sub FormPaper_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseHover
        Label11.ForeColor = Color.White
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label11.Text = "Click me for" & vbNewLine & "words of encouragement!"
    End Sub

    Public Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        Using g As Graphics = Graphics.FromImage(PictureBox1.Image)
            g.Clear(Color.White)
        End Using
        Dim tempID As Integer = Questions(Count)
        QuestionID = tempID
        PictureBox1.Invalidate()
        getQImage(PictureBox1)
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

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label2.Text = TimeOfDay.ToString("h:mm tt")
        If TimeOn = True Then
            Elapsed = Now - StartTime
            Label3.Text = "You're on the clock: " & Elapsed.Minutes & " minute(s) and " & Elapsed.Seconds & " second(s)!"
        End If
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Dim k As New Graphing.FormGraphs
        k.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MsgBox("Please note skipping questions may limit available test questions to display." & vbNewLine & "Are you sure you want to skip this question?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Skip Question") = MsgBoxResult.Yes Then
            Count = Count - 1
            MaxCount = Count
            Questions.Remove(Questions(Questions.Count - 1))
            Label8.Text = Convert.ToInt32(Label8.Text) - tempMark
            Tried = False
            GetQuestion()
        End If
    End Sub
End Class