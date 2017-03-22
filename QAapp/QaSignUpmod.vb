
Imports System.Data.OleDb

Module QaSignUpmod

    Public contemp As New OleDbConnection

    Public readertemp As OleDbDataReader

    Public sqltemp As String

    Dim Desk As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

    Sub connecttemp()



        Try

            contemp.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= P:\SPC\QA\QA.accdb"

            '   contemp.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"

            '  contemp.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp.Open()





        Catch ex As OleDbException



            MsgBox("Connection Break at 'QasignUpMod', please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try









    End Sub




End Module
