Public Class ExitForm
    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs) Handles ConfirmButton.Click
        RentalForm.Close()
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

    Private Sub QuestionLabel_Click(sender As Object, e As EventArgs) Handles QuestionLabel.Click

    End Sub

    Private Sub ExitForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class