Namespace My
    Partial Friend Class MyApplication

        Private Sub MyApplication_Startup(
            sender As Object,
            e As ApplicationServices.StartupEventArgs
        ) Handles Me.Startup
            UrlProtocol.RegisterQlnsProtocol()
        End Sub

    End Class
End Namespace