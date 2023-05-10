Public Class frmUpdateDeNghiKT
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
    Private _IdDonVi As Integer
    Private _CN_TT As Int16 '1: ca nhan; 3: Phong ban; 4: Don vi
    Private _IdDoiTuong As String
    Private _Nam As Integer
    Private _DinhKy As Int16

    Public Property IdDonVi() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal value As Integer)
            _IdDonVi = value
        End Set
    End Property

    Public Property CN_TT() As Int16
        Get
            Return _CN_TT
        End Get
        Set(ByVal value As Int16)
            _CN_TT = value
        End Set
    End Property

    Public Property IdDoiTuong() As String
        Get
            Return _IdDoiTuong
        End Get
        Set(ByVal value As String)
            _IdDoiTuong = value
        End Set
    End Property

    Public Property Nam() As Integer
        Get
            Return _Nam
        End Get
        Set(ByVal value As Integer)
            _Nam = value
        End Set
    End Property

    Public Property DinhKy() As Int16
        Get
            Return _DinhKy
        End Get
        Set(ByVal value As Int16)
            _DinhKy = value
        End Set
    End Property

    Private Sub frmUpdateDeNghiKT_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmUpdateDeNghiKT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bntClose.Visible = False
        lab.Text = "Năm " & Nam
        If CN_TT = 1 Then
            txtDoiTuong.Text = getFullName(IdDoiTuong, False)
        Else
            If CN_TT = 3 Then
                txtDoiTuong.Text = getPhong(IdDoiTuong) & getDonvi(IdDonVi)
            Else
                txtDoiTuong.Text = getDonvi(IdDoiTuong)
            End If
        End If
        initGrid(gridResult)
    End Sub

    Private Sub initGrid(ByRef gridName As DataGridView)
        gridName.Columns.Add("C1", "ID")
        gridName.Columns(0).Visible = False
        Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
        ckb_choice.Name = "cln_cbDuyet"
        ckb_choice.HeaderText = "Duyệt"
        ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        ckb_choice.Width = 50
        gridName.Columns.Add(ckb_choice)
        gridName.Columns.Add("C2", "Danh hiệu/Hình thức")
        'gridName.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        gridName.Columns(2).Width = 300
        gridName.Columns(2).ReadOnly = True
        gridName.Columns.Add("C4", "Nội dung")
        gridName.Columns(3).Width = 300
        gridName.Columns(3).ReadOnly = True
        Dim dt As DataTable = New DataTable
        dt = getDSDeNghiKT_CN_TT(Nam, KT_ChuyenMon, DinhKy, CN_TT, IdDoiTuong, IdDonVi)
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                gridName.Rows.Add()
                gridName.Rows(i).Cells("C1").Value = dt.Rows(i)("C1").ToString().Trim()
                gridName.Rows(i).Cells("C2").Value = dt.Rows(i)("C2").ToString().Trim()
                gridName.Rows(i).Cells("C4").Value = dt.Rows(i)("C4").ToString().Trim()
                If dt.Rows(i)("C3") = 1 Then
                    gridName.Rows(i).Cells("cln_cbDuyet").Value = True
                Else
                    gridName.Rows(i).Cells("cln_cbDuyet").Value = False
                End If
            Next

        End If
    End Sub

    Private Sub bntUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate.Click
        If (gridResult.Rows.Count > 0) Then
            For i As Int32 = 0 To gridResult.Rows.Count - 1
                Dim m_KhenThuongCT As New KhenThuong_CT
                m_KhenThuongCT.IdKhenThuong_CT = gridResult.Rows(i).Cells("C1").Value.ToString().Trim
                m_KhenThuongCT.DaDuyet = CType(gridResult.Rows(i).Cells("cln_cbDuyet").Value, Boolean)
                m_KhenThuongCT.UpdatenotDuyet()
            Next
        End If
        Close()
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub
End Class