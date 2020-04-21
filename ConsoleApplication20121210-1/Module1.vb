Module Module1

    Sub Main()

        Try
            Dim cls1 As New Class1()
            cls1.Done()
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
        End Try

    End Sub

End Module

