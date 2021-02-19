Public Class FormM
    Dim penselect As Boolean = False
    Dim MyPicture As Image = FormQandA.Picture

    Private Sub FormQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
        PictureBox1.Image = MyPicture
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
        PictureBox12_Click(Nothing, Nothing)
        GetMS()
        Me.Close()
    End Sub

    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        Using g As Graphics = Graphics.FromImage(PictureBox1.Image)
            g.Clear(Color.White)
        End Using
        PictureBox1.Invalidate()
        GetMS()
    End Sub

    Private Sub GetMS()
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
                    PictureBox1.Image = picture
                End Using
            End With
            cn.Close()
        End If
    End Sub
End Class