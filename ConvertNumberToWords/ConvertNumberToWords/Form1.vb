Public Class Form1
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        JustNumbers(TextBox1, e)
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        JustNumbers1(TextBox2, e)
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> "" Then
            Label1.Text = ConvertNumberToWords3(TextBox1.Text, TextBox3.Text, TextBox4.Text)
        Else
            Label1.Text = ""
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text <> "" Then
            Label2.Text = ConvertNumberToWords2(TextBox2.Text, TextBox3.Text, TextBox4.Text)
        Else
            Label2.Text = ""
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Text = "دينار"
        TextBox4.Text = "درهم"
    End Sub
End Class
