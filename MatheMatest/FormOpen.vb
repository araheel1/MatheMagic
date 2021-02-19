Public Class FormOpen
    Dim RanNum As New Random
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        FormSelectBook.Show()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        FormSelectBook.Show()
    End Sub

    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                PictureBox4.Image = My.Resources.download__2_1
            Case 2
                PictureBox4.Image = My.Resources.download__2_b
            Case 3
                PictureBox4.Image = My.Resources.download__2_o
            Case 4
                PictureBox4.Image = My.Resources.download__2_y
        End Select
    End Sub

    Private Sub Label3_MouseHover(sender As Object, e As EventArgs) Handles Label3.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label3.ForeColor = Color.Red
            Case 2
                Label3.ForeColor = Color.Blue
            Case 3
                Label3.ForeColor = Color.Orange
            Case 4
                Label3.ForeColor = Color.Yellow
        End Select
    End Sub


    Private Sub Label2_MouseHover(sender As Object, e As EventArgs) Handles Label2.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label2.ForeColor = Color.Red
            Case 2
                Label2.ForeColor = Color.Blue
            Case 3
                Label2.ForeColor = Color.Orange
            Case 4
                Label2.ForeColor = Color.Yellow
        End Select
    End Sub

    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                PictureBox2.Image = My.Resources.download__1_1
            Case 2
                PictureBox2.Image = My.Resources.download__1_b
            Case 3
                PictureBox2.Image = My.Resources.download__1_o
            Case 4
                PictureBox2.Image = My.Resources.download__1_y
        End Select
    End Sub

    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs) Handles PictureBox3.MouseHover
        Label3.ForeColor = Color.White
        Label2.ForeColor = Color.White
        Label4.ForeColor = Color.White
        Label1.ForeColor = Color.White
        PictureBox2.Image = My.Resources.download__1_
        PictureBox4.Image = My.Resources.download__2_
    End Sub

    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label1.ForeColor = Color.Red
            Case 2
                Label1.ForeColor = Color.Blue
            Case 3
                Label1.ForeColor = Color.Orange
            Case 4
                Label1.ForeColor = Color.Yellow
        End Select
    End Sub

    Private Sub Label4_MouseHover(sender As Object, e As EventArgs) Handles Label4.MouseHover
        Dim LowerBound As Integer = 1
        Dim UpperBound As Integer = 5
        Dim Selection As Integer
        Selection = RanNum.Next(LowerBound, UpperBound)
        Select Case Selection
            Case 1
                Label4.ForeColor = Color.Red
            Case 2
                Label4.ForeColor = Color.Blue
            Case 3
                Label4.ForeColor = Color.Orange
            Case 4
                Label4.ForeColor = Color.Yellow
        End Select
    End Sub

    Private Sub FormOpen_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class