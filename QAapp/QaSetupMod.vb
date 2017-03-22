'
Imports System.Data.OleDb
Imports System.Data.SqlClient

Module QaSetupMod



    Public contemp1 As New SqlConnection
    Public contemp2 As New SqlConnection
    Public contemp3 As New SqlConnection
    Public contemp4 As New SqlConnection
    Public contemp5 As New SqlConnection
    Public contemp6 As New SqlConnection
    Public contemp7 As New SqlConnection
    Public contemp8 As New SqlConnection
    Public contemp9 As New SqlConnection
    Public contemp10 As New SqlConnection
    Public contemp11 As New SqlConnection
    Public contemp12 As New SqlConnection
    Public contemp13 As New SqlConnection
    Public contemp14 As New SqlConnection
    Public contemp15 As New SqlConnection
    Public contemp16 As New SqlConnection
    Public contemp17 As New SqlConnection
    Public contemp18 As New SqlConnection
    Public contemp19 As New SqlConnection
    Public contemp20 As New SqlConnection
    Public contemp21 As New SqlConnection

    Public contemp1a As New SqlConnection
    Public contemp1b As New SqlConnection
    Public contemp1c As New SqlConnection




    Public contemp00 As New SqlConnection

    Public readertemp1 As SqlDataReader
    Public readertemp2 As SqlDataReader
    Public readertemp3 As SqlDataReader
    Public readertemp4 As SqlDataReader
    Public readertemp5 As SqlDataReader
    Public readertemp6 As SqlDataReader
    Public readertemp7 As SqlDataReader
    Public readertemp8 As SqlDataReader
    Public readertemp9 As SqlDataReader
    Public readertemp10 As SqlDataReader
    Public readertemp11 As SqlDataReader
    Public readertemp12 As SqlDataReader
    Public readertemp13 As SqlDataReader
    Public readertemp14 As SqlDataReader
    Public readertemp15 As SqlDataReader
    Public readertemp16 As SqlDataReader
    Public readertemp17 As SqlDataReader
    Public readertemp18 As SqlDataReader
    Public readertemp19 As SqlDataReader
    Public readertemp20 As SqlDataReader
    Public readertemp21 As SqlDataReader

    Public readertemp00 As SqlDataReader



    Public readertemp1a As SqlDataReader
    Public readertemp1b As SqlDataReader
    Public readertemp1c As SqlDataReader

    Public sqltemp1a As String
    Public sqltemp1b As String
    Public sqltemp1c As String

    Public sqltemp1 As String
    Public sqltemp2 As String
    Public sqltemp3 As String
    Public sqltemp4 As String
    Public sqltemp5 As String
    Public sqltemp6 As String
    Public sqltemp7 As String
    Public sqltemp8 As String
    Public sqltemp9 As String
    Public sqltemp10 As String
    Public sqltemp11 As String
    Public sqltemp12 As String
    Public sqltemp13 As String
    Public sqltemp14 As String
    Public sqltemp15 As String
    Public sqltemp16 As String
    Public sqltemp17 As String
    Public sqltemp18 As String
    Public sqltemp19 As String
    Public sqltemp20 As String

    Public sqltemp00 As String

    Public sqltemp21 As String


    '  Dim Desk As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

    Dim Desk = My.Computer.FileSystem.SpecialDirectories.Desktop

    Sub connecttemp1()



        Try

            contemp1.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp1.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 1, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub



    Sub connecttemp2()



        Try

            contemp2.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '   contemp2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp2.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 2, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub

    Sub connecttemp3()



        Try

            contemp3.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp3.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"


            '  contemp3.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp3.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 3, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub

    Sub connecttemp4()



        Try

            contemp4.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp4.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"



            '  contemp4.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp4.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 4, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub

    Sub connecttemp5()



        Try

            contemp5.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '   contemp5.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"



            '  contemp5.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp5.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 5, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub

    Sub connecttemp6()



        Try

            contemp6.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '    contemp6.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"



            ' contemp6.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp6.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 6, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub





    Sub connecttemp7()



        Try

            contemp7.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp7.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"



            '  contemp7.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp7.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 7, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub



    Sub connecttemp8()



        Try

            contemp8.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '   contemp8.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp8.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"


            contemp8.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 8, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try




    End Sub



    Sub connecttemp9()



        Try

            contemp9.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp9.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"


            '  contemp9.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp9.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 9, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub

    Sub connecttemp10()



        Try

            contemp10.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp10.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"


            ' contemp10.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp10.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 10, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub




    Sub connecttemp11()



        Try

            contemp11.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp11.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"

            '  contemp11.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp11.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 11, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub


    Sub connecttemp00()



        Try

            contemp00.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp00.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QA.accdb"

            '  contemp00.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp00.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 00, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub



    Sub connecttemp13()



        Try

            contemp13.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp13.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp12.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp13.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 13, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub


    Sub connecttemp12()



        Try

            contemp12.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp12.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp12.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp12.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 12, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try

    End Sub


    Sub connecttemp14()



        Try

            contemp14.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp14.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp14.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp14.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 14, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try



    End Sub


    Sub connecttemp15()



        Try

            contemp15.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '   contemp15.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp15.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp15.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 15, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try



    End Sub




    Sub connecttemp1a()



        Try

            contemp1a.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '    contemp1a.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp12.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp1a.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 1a, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try



    End Sub



    Sub connecttemp1b()



        Try

            contemp1b.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp1b.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp1b.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp1b.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 1b, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try



    End Sub


    Sub connecttemp1c()



        Try

            contemp1c.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            '  contemp1c.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '  contemp1c.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"

            contemp1c.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 1c, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try



    End Sub



    Sub connecttemp16()



        Try

            contemp16.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp16.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 16, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub


    Sub connecttemp17()



        Try

            contemp17.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp17.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 17, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub




    Sub connecttemp18()



        Try

            contemp18.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp18.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 18, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub


    Sub connecttemp19()



        Try

            contemp19.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp19.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 19, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub




    Sub connecttemp20()



        Try

            contemp20.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp20.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 20, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub


    Sub connecttemp21()



        Try

            contemp21.ConnectionString = "Server=tcp:edurrant.database.windows.net,1433;Initial Catalog=QADB;Persist Security Info=False;User ID=playergoodi;Password=Grinder3$; MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

            ' contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\playe\Desktop\QA\QABackEnd.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lblSCRN.text & "\Desktop\QA1\QA.accdb"

            '   contemp1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Desk & "\QA1\QA.accdb"




            contemp21.Open()





        Catch ex As SqlException



            MsgBox("Connection Break at 'connectTemp 21, please restart")



            MsgBox(ex.Message)



        Catch ex As SystemException



            MsgBox(ex.Message)



        Catch ex As Exception



            MsgBox(ex.Message)





        End Try







    End Sub

End Module
