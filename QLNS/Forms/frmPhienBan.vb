Public Class frmPhienBan
    Private Sub frmPhienBan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'tbx_LoadTextFile(txtPhienBan, "versions.txt")
        'txtPhienBan.Select(0, 0)
        Try
            rtb.LoadFile("versions.rtf")
        Catch ex As Exception
            My_MessageBox("Không đọc được nội dung phiên bản!")
        End Try
    End Sub

    Private Sub bntClose_Click(sender As Object, e As EventArgs) Handles bntClose.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class