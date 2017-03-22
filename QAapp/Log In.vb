

Imports System.Data.OleDb
Imports System.Data.SqlClient

Module Log_In


    Public conlogin As New SqlConnection

    Public readerlogin As SqlDataReader

    Public sqllogin As String


    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop

    Sub connectlogin()

        Try


            '   contemp8.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            '   conlogin.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"


            conlogin.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"






            '   conlogin.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            '  conlogin.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & QaSetup.lblMDrive.Text & "\Users\" & QaSetup.lblSCRN.Text & "\Desktop\QA1\QA.accdb"

            conlogin.Open()




        Catch ex As OleDbException

            MsgBox(ex.Message)

        Catch ex As SystemException

            MsgBox(ex.Message)

        Catch ex As Exception

            MsgBox(ex.Message)


        End Try




    End Sub

















End Module
