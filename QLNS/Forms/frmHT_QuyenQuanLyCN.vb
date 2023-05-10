Public Class frmHT_QuyenQuanLyCN
#Region "Form events"
    Sub New(IdCanBo As String, HoTen As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtHoTen.Text = HoTen
        txtMaCB.Text = IdCanBo
    End Sub

    Private Sub frmHT_QuyenQuanLyCN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _LoadDSChiNhanh()
        _LoadQuyenQuanLyCN()
    End Sub

    Private Sub btn_accept_Click(sender As Object, e As EventArgs) Handles btn_accept.Click
        Dim sRights As String = chkl_GetItemsCheck(chklChiNhanh)
        Dim conn As New DBAccess
        conn.executeSQL("DELETE FROM dbo.HT_QuyenQuanLyCN WHERE IdCanBo = '" & txtMaCB.Text & "'")

        'Nếu check toàn bộ đơn vị thì không cần đưa vào bản Quyền quản lý nữa
        If sRights.Split(";").Length < chklChiNhanh.Items.Count Then
            conn.executeSQL("INSERT INTO dbo.HT_QuyenQuanLyCN(IdCanBo, QuyenQuanLyCN) VALUES ('" & txtMaCB.Text & "','" & sRights & "')")
        End If
        Me.Close()
    End Sub

    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        Me.Close()
    End Sub

    Private Sub chkCheckAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkCheckAll.CheckedChanged
        With chklChiNhanh
            If .Items.Count > 0 Then
                For i As Integer = 0 To .Items.Count - 1
                    .SetItemChecked(i, chkCheckAll.Checked)
                Next
            End If
        End With
    End Sub
#End Region

#Region "Local functions"
    Sub _LoadDSChiNhanh()
        'SELECT Id, ten_goi FROM ChiNhanh WHERE id_goc = 1 And [Status]=1 Order By id
        Dim conn As New DBAccess
        'Lấy danh sách
        chkl_Binding(chklChiNhanh, conn.getDataTable("SELECT Id, ten_goi FROM ChiNhanh WHERE id_goc <= 1 And [Status]=1 Order By id"))
    End Sub

    Sub _LoadQuyenQuanLyCN()
        Dim conn As New DBAccess
        Dim sRights As String = conn.getString("SELECT QuyenQuanLyCN FROM dbo.HT_QuyenQuanLyCN WHERE IdCanBo = '" & txtMaCB.Text & "'")
        If (Not sRights Is Nothing) AndAlso sRights.Length > 0 Then
            chkl_SetItemsCheck(chklChiNhanh, sRights)
        Else
            chkCheckAll.Checked = True
        End If
    End Sub
#End Region

End Class