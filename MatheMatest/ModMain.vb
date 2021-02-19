Imports System.Data.OleDb
Module ModMain
    Public Const DataBasePath As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Users\Khalid\Desktop\FMAnew.mdb';Persist Security Info=False;"
    'Public DataBasePath As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Application.StartupPath & "\FMAnew.mdb';Persist Security Info=False;"
    Public cn As OleDbConnection
    Public QuestionID, TopicID, ModuleID As Integer
    Public Name As String
    Public itemList As New List(Of Integer)
    Public Questions As New List(Of Integer)
    Public PubAtt As Integer

    Public Function DbConnect() As Boolean
        Try
            cn = New OleDbConnection(DataBasePath)
            cn.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to open the database :" & ex.Message)
            Return False
        End Try
    End Function

    Public Class ListBoxData
        Public Data As String
        Public Identifier As Integer

        Public Sub New(NewData As String, NewIdentifier As Integer)
            Data = NewData
            Identifier = NewIdentifier
        End Sub

        Public Overrides Function ToString() As String
            Return Data
        End Function
    End Class

    Public Sub getQImage(pbx As PictureBox)
        If DbConnect() Then
            Dim SQLCmd As New OleDb.OleDbCommand
            With SQLCmd

                .Connection = cn
                .CommandText = "SELECT QImage FROM Questions WHERE QuestionID = @ID"
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
End Module
