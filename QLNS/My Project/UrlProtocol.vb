Imports Microsoft.Win32
Imports System.Diagnostics

Public Module UrlProtocol
    Public Sub RegisterQlnsProtocol()
        Dim exePath = Process.GetCurrentProcess().MainModule.FileName

        Using key = Registry.CurrentUser.CreateSubKey(
            "Software\Classes\qlns\shell\open\command")
            key.SetValue("", """" & exePath & """ ""%1""")
        End Using

        Using key = Registry.CurrentUser.CreateSubKey("Software\Classes\qlns")
            key.SetValue("", "URL:QLNS Protocol")
            key.SetValue("URL Protocol", "")
        End Using
    End Sub
End Module