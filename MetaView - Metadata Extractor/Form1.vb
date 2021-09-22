Imports System.IO

Public Class Form1

    Dim path As String
    Dim tool As String
    Dim Target() As String
    Dim url As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox2.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            PathBox.Text = OpenFileDialog1.SafeFileName
            path = OpenFileDialog1.FileName
            tool = My.Computer.FileSystem.SpecialDirectories.Temp + "exiftool.exe"
            IO.File.WriteAllBytes(tool, My.Resources.exiftool)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        PictureBox2.Visible = False

        If path = "" Then
            MessageBox.Show("File Not Found")
        Else
            RichTextBox1.SelectionStart = 0
            RichTextBox1.SelectionLength = 0

            If path.Contains(".jpg") = True Or path.Contains(".png") = True Then
                PictureBox1.ImageLocation = path
            ElseIf path.Contains(".mp4") = True Or path.Contains(".avi") = True Or path.Contains(".mkv") = True Then
                PictureBox1.Image = My.Resources.video_icon
            Else
                PictureBox1.Image = My.Resources.file_icon
            End If

            RichTextBox1.Text = ""

            Dim p = New System.Diagnostics.Process()
            p.StartInfo.FileName = tool
            p.StartInfo.Arguments = """" + path + """"
            p.StartInfo.UseShellExecute = False
            p.StartInfo.CreateNoWindow = True
            p.StartInfo.RedirectStandardOutput = True
            p.Start()

            RichTextBox1.Text = p.StandardOutput.ReadToEnd()

            If RichTextBox1.Text.Contains("GPS Position") Then
                GPS()
            End If

        End If
    End Sub

    Sub GPS()

        PictureBox2.Visible = True

        Dim T1 As String = "GPS Position                    : "

        Dim pFrom As Integer = RichTextBox1.Text.IndexOf(T1) + T1.Length
        Dim position As String = RichTextBox1.Text.Substring(pFrom, 40)
        'MessageBox.Show(position)
        position = position.Replace(" deg ", " ")
        position = position.Replace(", ", " ")
        position = position.Replace("' ", " ")
        position = position.Replace(""" ", " ")

        'MessageBox.Show(position)

        Dim ch As Char
        Dim count As Integer = 0
        Dim d1, m1, s1, d2, m2, s2 As Double
        Dim pole1, pole2 As Char
        Dim latitude, longitude As Double

        For Each ch In position
            If Convert.ToInt32(ch) >= 48 And Convert.ToInt32(ch) <= 57 Or ch = "." Or ch = "N" Or ch = "E" Then
                Select Case count
                    Case 0
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        d1 = Convert.ToDouble(temp)
                    Case 1
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        m1 = Convert.ToDouble(temp)
                    Case 2
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        s1 = Convert.ToDouble(temp)
                    Case 3
                        pole1 = ch
                    Case 4
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        d2 = Convert.ToDouble(temp)
                    Case 5
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        m2 = Convert.ToDouble(temp)
                    Case 6
                        Dim temp As String
                        temp = temp + ch.ToString(ch)
                        s2 = Convert.ToDouble(temp)
                    Case 7
                        pole2 = ch
                End Select
            Else
                count = count + 1
            End If
        Next

        latitude = d1 + (m1 / 60) + (s1 / 3600)
        longitude = d2 + (m2 / 60) + (s2 / 3600)

        'MessageBox.Show(latitude)
        'MessageBox.Show(longitude)

        Dim glink As String = "www.google.com/maps/search/?api=1&query=" + latitude.ToString() + "%2C" + longitude.ToString()

        RichTextBox1.Text = RichTextBox1.Text + "Google Maps Link       : " + glink

        url = glink

    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Process.Start(url)
    End Sub


End Class
