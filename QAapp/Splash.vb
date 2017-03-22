Public Class Splash

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer222.Tick



        ProgressBar1.PerformStep()



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer333.Tick


   



    End Sub



    Private Sub Splash_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Me.CenterToScreen()



        Timer333.Start()
        Timer222.Start()



    End Sub


End Class