Imports DankBarcodeUtils
Public Class Form1

    Dim saveble As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        generate()
    End Sub

    Sub generate()
        If ComboBox1.Text = "EAN8" Then
            PictureBox1.Image = New Barcode.Ean8(TextBox1.Text, CheckBox1.Checked, CheckBox2.Checked).GetImage(TrackBar1.Value, 90, CheckBox3.Checked)
            saveble = True
        ElseIf ComboBox1.Text = "EAN13" Then
            PictureBox1.Image = New Barcode.Ean13(TextBox1.Text, CheckBox1.Checked, CheckBox2.Checked).GetImage(TrackBar1.Value, 90, CheckBox3.Checked)
            saveble = True
        ElseIf ComboBox1.Text = "CODE128" Then
            PictureBox1.Image = New Barcode.Code128(TextBox1.Text).GetImage(TrackBar1.Value, 90, CheckBox3.Checked)
            saveble = True
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "EAN8" Then
            TextBox1.Enabled = True
            TextBox1.MaxLength = 7
        ElseIf ComboBox1.Text = "EAN13" Then
            TextBox1.Enabled = True
            TextBox1.MaxLength = 12
        ElseIf ComboBox1.Text = "CODE128" Then
            TextBox1.Enabled = True
            TextBox1.MaxLength = 128
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If saveble = True Then
            Dim r As DialogResult = SaveFileDialog1.ShowDialog
            If r = Windows.Forms.DialogResult.OK Then
                Try
                    PictureBox1.Image.Save(SaveFileDialog1.FileName)
                Catch ex As Exception
                    MessageBox.Show("Error while saving barcode", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        generate()
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        generate()
    End Sub
End Class
