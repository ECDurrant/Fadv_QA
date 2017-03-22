Public Class QADash
    Private Sub QADash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'QADataSet.QAMainDB' table. You can move, or remove it, as needed.
        Me.QAMainDBTableAdapter.Fill(Me.QADataSet.QAMainDB)

    End Sub
End Class