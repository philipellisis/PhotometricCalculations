Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim spdText As String() = TextBox1.Text.Split(vbLf)
            Dim spd() As Double
            ReDim spd(spdText.Length - 1)
            For i As Integer = 0 To spdText.Length - 1
                spd(i) = CDbl(spdText(i))
            Next
            Dim cct As New CCT
            Dim cri As New CRI
            cct.calculate(spd)
            cri.calculate(spd, cct.cct)
            lblCRI.Text = cri.CRI.ToString("F1")
            lblCCT.Text = cct.cct.ToString("F0")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class
