Imports System.Drawing.Imaging

Public Class FormSelectBook


    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover
        Label1.ForeColor = System.Drawing.Color.Firebrick
    End Sub

    Private Sub Label2_MouseHover(sender As Object, e As EventArgs) Handles Label2.MouseHover
        Label2.ForeColor = System.Drawing.Color.Firebrick
    End Sub

    Private Sub Label3_MouseHover(sender As Object, e As EventArgs) Handles Label3.MouseHover
        Label3.ForeColor = System.Drawing.Color.Firebrick
    End Sub

    Private Sub Label4_MouseHover(sender As Object, e As EventArgs) Handles Label4.MouseHover
        Label4.ForeColor = System.Drawing.Color.Firebrick
    End Sub

    Private Sub Label5_MouseHover(sender As Object, e As EventArgs) Handles Label5.MouseHover
        Label5.ForeColor = System.Drawing.Color.Firebrick
    End Sub


    Private Sub PictureBox7_MouseHover1(sender As Object, e As EventArgs) Handles PictureBox7.MouseHover
        Label1.ForeColor = System.Drawing.Color.Blue
        Label2.ForeColor = System.Drawing.Color.Blue
        Label3.ForeColor = System.Drawing.Color.Blue
        Label4.ForeColor = System.Drawing.Color.Blue
        Label5.ForeColor = System.Drawing.Color.Blue
        PictureBox1.Image = My.Resources.C11
        PictureBox2.Image = My.Resources.C21
        PictureBox5.Image = My.Resources.FP1
        PictureBox3.Image = My.Resources.FM1
        PictureBox4.Image = My.Resources.D1
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        PictureBox1.Image = My.Resources.c1b
    End Sub


    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        PictureBox2.Image = My.Resources.c2b
    End Sub

    Private Sub PictureBox5_MouseHover(sender As Object, e As EventArgs) Handles PictureBox5.MouseHover
        PictureBox5.Image = My.Resources.fp1b
    End Sub



    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs) Handles PictureBox3.MouseHover
        PictureBox3.Image = My.Resources.fm1b
    End Sub



    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        PictureBox4.Image = My.Resources.d1b
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        ModuleID = 1
        FormC1.ShowDialog()
    End Sub


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        ModuleID = 1
        FormC1.ShowDialog()
    End Sub


    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        ModuleID = 2
        FormC2.ShowDialog()
    End Sub


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        ModuleID = 2
        FormC2.ShowDialog()
    End Sub


    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        ModuleID = 3
        FormFP1.ShowDialog()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        ModuleID = 3
        FormFP1.ShowDialog()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        ModuleID = 4
        FormFM1.ShowDialog()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        ModuleID = 4
        FormFM1.ShowDialog()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        FormD1.ShowDialog()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        MsgBox("'Bantz' lmao wyd fam")
    End Sub

    Private Sub PictureBox6_MouseHover(sender As Object, e As EventArgs) Handles PictureBox6.MouseHover
        Label1.ForeColor = System.Drawing.Color.Blue
        PictureBox1.Image = My.Resources.C11
    End Sub

    Private Sub PictureBox12_MouseHover(sender As Object, e As EventArgs) Handles PictureBox12.MouseHover
        Label2.ForeColor = System.Drawing.Color.Blue
        PictureBox2.Image = My.Resources.C21
    End Sub

    Private Sub PictureBox8_MouseHover(sender As Object, e As EventArgs) Handles PictureBox8.MouseHover
        Label3.ForeColor = System.Drawing.Color.Blue
        PictureBox5.Image = My.Resources.FP1
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        Label4.ForeColor = System.Drawing.Color.Blue
        PictureBox3.Image = My.Resources.FM1
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        Label5.ForeColor = System.Drawing.Color.Blue
        PictureBox4.Image = My.Resources.D1
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Name = InputBox("Please enter your name: ")
        Label6.Text = "Welcome, " & Name
    End Sub

End Class