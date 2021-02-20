Public Class Form2

    Private Sub btnnew_Click(sender As Object, e As EventArgs) Handles btnnew.Click
        Form3.Show()
        Me.Hide()
    End Sub

  
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = Form1.txtun.Text()
    End Sub
End Class