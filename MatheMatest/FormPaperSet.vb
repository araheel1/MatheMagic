Public Class FormPaperSet

    Private Sub FormPaperSet_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetTopics()
        itemList.Clear()
        Colors()
    End Sub

    Private Sub Colors()
        If (ModuleID = 1) Or (ModuleID = 2) Then
            PictureBox1.BackColor = Color.ForestGreen
            Label4.BackColor = Color.ForestGreen
            Label4.ForeColor = Color.White
            Me.BackColor = Color.PaleGreen
            Label3.BackColor = Color.PaleGreen
            Label3.ForeColor = Color.Black
            Label2.BackColor = Color.ForestGreen
            Label2.ForeColor = Color.White
            Label1.BackColor = Color.ForestGreen
            Label1.ForeColor = Color.White
            TrackBar1.BackColor = Color.ForestGreen
            TrackBar1.ForeColor = Color.White
            ListBox1.BackColor = Color.ForestGreen
            ListBox1.ForeColor = Color.White
            Button3.BackColor = Color.ForestGreen
            Button3.ForeColor = Color.White
            PictureBox2.Image = My.Resources._154910634887195497
            PictureBox4.BackColor = Color.ForestGreen
            PictureBox3.BackColor = Color.ForestGreen

        ElseIf ModuleID = 3 Then
            PictureBox1.BackColor = Color.DarkCyan
            Label4.BackColor = Color.DarkCyan
            Label4.ForeColor = Color.White
            Me.BackColor = Color.CadetBlue
            Label3.BackColor = Color.CadetBlue
            Label3.ForeColor = Color.White
            Label2.BackColor = Color.DarkCyan
            Label2.ForeColor = Color.White
            Label1.BackColor = Color.DarkCyan
            Label1.ForeColor = Color.White
            TrackBar1.BackColor = Color.DarkCyan
            TrackBar1.ForeColor = Color.White
            ListBox1.BackColor = Color.DarkCyan
            ListBox1.ForeColor = Color.White
            Button3.BackColor = Color.DarkCyan
            Button3.ForeColor = Color.White
            PictureBox2.Image = My.Resources.fp1boat
            PictureBox4.BackColor = Color.DarkCyan
            PictureBox3.BackColor = Color.DarkCyan
        End If
        If ModuleID = 1 Then
            Label4.Text = "Core Pure Mathematics: Book 1/AS" & vbNewLine & "Paper Creator"
        ElseIf ModuleID = 2 Then
            Label4.Text = "Core Pure Mathematics: Book 2" & vbNewLine & "Paper Creator"
        Else
            Label4.Text = "Further Pure Mathematics 1" & vbNewLine & "Paper Creator"
        End If
    End Sub

    Private Sub GetTopics()
        If DbConnect() Then
            LstSelectCust.Items.Clear()
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd
                .Connection = cn
                .CommandText = "Select * From Topic Where ModuleID = @MID"
                .Parameters.AddWithValue("@MID", ModuleID)
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
            ListBox1.Items.Add(SelectedCust)
            LstSelectCust.Items.Remove(SelectedCust)
            itemList.Add(SelectedCust.Identifier)
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        ListBox1.Items.Clear()
        itemList.Clear()
        Questions.Clear()
        TextBox1.Text = 50
        TrackBar1.Value = 50
        GetTopics()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.Items.Count = 0 Then
            MsgBox("No topics selected!")
        Else
            ListBox1.SelectedIndex = 0
            Dim f2 As New FormPaper
            f2.marks = TrackBar1.Value
            f2.Show()
        End If
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        TextBox1.Text = TrackBar1.Value
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim rnd As New Random
        Dim randomIndex As Integer = rnd.Next(0, LstSelectCust.Items.Count - 1)
        LstSelectCust.SelectedItem = (LstSelectCust.Items(randomIndex))
    End Sub
End Class