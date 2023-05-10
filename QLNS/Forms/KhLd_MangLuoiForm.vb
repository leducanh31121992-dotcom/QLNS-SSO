Imports System
Imports System.Data
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient

Public Class KhLd_MangLuoiForm
    Private vNFInfo_D0 As System.Globalization.NumberFormatInfo
    Private vNFInfo_D2 As System.Globalization.NumberFormatInfo
    Private _KHLD_MangLuoiBLL As KHLD_MangLuoiBLL = New KHLD_MangLuoiBLL()
    Private _LoginPos As String = ""
    Private _LevelRun As String = ""
    Private _FlagTWCN As Byte = 0
    Private _FlagAdd As Boolean = False
    Private _Globals As Globals = New Globals
    Private ARL_KhLd_DonVi As ArrayList = New ArrayList
    Private ARL_NhuCau_DonVi As ArrayList = New ArrayList
    Private IdKHML_Sele As String = ""

    Dim _ClassNumber As Integer = 3

    Public Sub New()
        _LoginPos = SoftSqlHelper.GetString(String.Format("Select Login_POS From HS_CanBo Where Login_Username='{0}'", gUsername), "")
        Dim _IdGoc As Integer = SoftSqlHelper.GetNumber(String.Format("Select Id_Goc From ChiNhanh Where Ma_So='{0}'", _LoginPos), 0)
        _LevelRun = IIf(_LoginPos = gMaDonViTW, 1, IIf(_IdGoc = 0 Or _IdGoc = 1, 2, 3))
        _FlagTWCN = _LevelRun

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo_D2 = New System.Globalization.NumberFormatInfo()
        vNFInfo_D2.NumberDecimalDigits = 2
        vNFInfo_D2.NumberGroupSeparator = ","

        vNFInfo_D0 = New System.Globalization.NumberFormatInfo()
        vNFInfo_D0.NumberDecimalDigits = 0
        vNFInfo_D0.NumberGroupSeparator = ","
    End Sub

#Region "---> Functions Main <---"
    Private Sub SetStyleRowGrid(ByVal iRow As Integer, ByVal dgv_Name As DataGridView, ByVal iKieuIn As Byte)
        dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Navy
        If iKieuIn = 0 Or iKieuIn = 1 Or iKieuIn = Nothing Then
            dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
            dgv_Name.Rows(iRow).DefaultCellStyle.BackColor = System.Drawing.Color.Orange
            dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = System.Drawing.Color.GhostWhite
        Else
            If iKieuIn = 4 Then
                dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.OrangeRed
            Else
                If iKieuIn = 2 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                Else
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="_DonVi_Cd"></param>
    ''' <param name="_ThoiDiem"></param>
    ''' <param name="_Loai_DL"></param>
    ''' <param name="_SoLDDH"></param>
    ''' <param name="_SoLDNH"></param>
    ''' <param name="_IdKHML"></param>
    ''' <param name="_FlagCall">Chỉ số xác định gọi hàm: 
    '''                                 2 - Kế hoạch lao động do TW nhập cho các Chi nhánh;
    '''                                 3 - Kế hoạch lao động do Chi nhánh nhập cho các đơn vị trực thuộc;
    '''                                 4 - Nhu cầu tuyển dụng lao động và Số cán bộ được thông báo trúng tuyển vào chi nhánh trong tháng đang tham gia lớp đào tạo.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValid(ByVal _DonVi_Cd As String, ByVal _ThoiDiem As DateTime, ByVal _Loai_DL As Byte, ByVal _SoLDDH As Integer, ByVal _SoLDNH As Integer, ByVal _NhuCauBS_SL As Integer, ByVal _ThBaoTTDangDT_SL As Integer, ByVal _IdKHML As String, ByVal _FlagCall As Byte) As Boolean
        Dim sIdKHML As String = "", sSQLCheck As String = ""

        If _FlagCall = 2 Or _FlagCall = 3 Then
            If _SoLDDH + _SoLDNH <= 0 Then
                MessageBox.Show("Bạn chưa nhập Kế hoạch lao động Dài hạn/Ngắn hạn. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_khld_sodaihan
                Return False
            End If

            If String.IsNullOrEmpty(_IdKHML) Then
                sIdKHML = SoftSqlHelper.GetString(String.Format("Select IsNull(IdKHML,'') From KHLD_MangLuoi Where ThoiDiem='{0}' And Loai_DL={1} And DonVi_Cd='{2}'", _ThoiDiem.ToString("yyyy-MM-dd"), _Loai_DL, _DonVi_Cd), "")
            Else
                sIdKHML = SoftSqlHelper.GetString(String.Format("Select IsNull(IdKHML,'') From KHLD_MangLuoi Where ThoiDiem='{0}' And Loai_DL={1} And DonVi_Cd='{2}' And IdKHML <> '{3}'", _ThoiDiem.ToString("yyyy-MM-dd"), _Loai_DL, _DonVi_Cd, _IdKHML), "")
            End If
            If Not String.IsNullOrEmpty(sIdKHML) Then
                MessageBox.Show("KHLĐ của đơn vị '" & cb_khld_donvi_tb.SelectedItem.ToString().Trim() & "' đã tồn tại. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_khld_donvi_tb
                Return False
            End If

            'Kiểm tra đối với Trường hợp Cập nhật KHLĐ của Chi nhánh Cho các PGD => Khi đó phải kiểm tra trước đó Chi nhánh đã được TW thông báo chưa?
            sIdKHML = ""
            If _Loai_DL = 3 Then
                If _DonVi_Cd <> _LoginPos Then
                    sSQLCheck = String.Format("Select IsNull(IdKHML,'') From KHLD_MangLuoi Where Loai_DL = 2 And ThoiDiem <= '{0}' And DonVi_Cd In (Select Ma_So From ChiNhanh Where Id In (Select Id_Goc From ChiNhanh Where Ma_So='{1}'))", _ThoiDiem.ToString("yyyy-MM-dd"), _DonVi_Cd)
                Else
                    sSQLCheck = String.Format("Select * From KHLD_MangLuoi Where Loai_DL = 2 And ThoiDiem <= '{0}' And DonVi_Cd IN (Select Ma_So From ChiNhanh Where Id In (Select Id From ChiNhanh Where Ma_So='{1}'))", _ThoiDiem.ToString("yyyy-MM-dd"), _DonVi_Cd)
                End If
                sIdKHML = SoftSqlHelper.GetString(sSQLCheck, "")
                If String.IsNullOrEmpty(sIdKHML) Then
                    MessageBox.Show("KHLĐ của Chi nhánh thời điểm '" & Format(CType(_ThoiDiem, DateTime), "dd-MM-yyyy") & "' chưa được TW phân giao. Do đó Chi nhánh không thể thông báo cho các đơn vị PGD trực thuộc. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_khld_ngaytb
                    Return False
                End If
            End If
        End If

        If _FlagCall = 4 Then
            If _NhuCauBS_SL + _ThBaoTTDangDT_SL <= 0 Then
                MessageBox.Show("Bạn chưa nhập [Nhu cầu lao động cần bổ sung/Cán bộ được thông báo trúng tuyển vào chi nhánh trong tháng đang tham gia lớp đào tạo]. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                If _NhuCauBS_SL <= 0 Then
                    ActiveControl = edt_nhucau_nhucaubs_sl
                Else
                    ActiveControl = edt_nhucau_thbaottdangdt_sl
                End If
                Return False
            End If

            If String.IsNullOrEmpty(_IdKHML) Then
                sIdKHML = SoftSqlHelper.GetString(String.Format("Select IsNull(IdKHML,'') From KHLD_MangLuoi Where ThoiDiem='{0}' And Loai_DL={1} And DonVi_Cd='{2}'", _ThoiDiem.ToString("yyyy-MM-dd"), _Loai_DL, _DonVi_Cd), "")
            Else
                sIdKHML = SoftSqlHelper.GetString(String.Format("Select IsNull(IdKHML,'') From KHLD_MangLuoi Where ThoiDiem='{0}' And Loai_DL={1} And DonVi_Cd='{2}' And IdKHML <> '{3}'", _ThoiDiem.ToString("yyyy-MM-dd"), _Loai_DL, _DonVi_Cd, _IdKHML), "")
            End If
            If Not String.IsNullOrEmpty(sIdKHML) Then
                MessageBox.Show("Nhu cầu lđ bổ sung - Thông báo trúng tuyển vào trong tháng đang tham gia đào tạo của đơn vị '" & cb_nhucau_donvi.SelectedItem.ToString().Trim() & "' đã tồn tại. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_nhucau_donvi
                Return False
            End If
        End If
        Return True
    End Function
#End Region

#Region "---> Functions Merget DataGridView - Nhu Cầu (Mới) <---"
    Private Sub dgv_nhucau_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            Dim r2 As Rectangle = e.CellBounds
            r2.Y += e.CellBounds.Height / 3
            r2.Height = e.CellBounds.Height / 3
            e.PaintBackground(r2, True)
            e.PaintContent(r2)
            e.Handled = True
        End If
    End Sub

    Private Sub dgv_nhucau_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_nhucau.DisplayRectangle
        rtHeader.Height = Me.dgv_nhucau.ColumnHeadersHeight / 2
        Me.dgv_nhucau.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_nhucau_ColumnWidthChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_nhucau.DisplayRectangle
        rtHeader.Height = Me.dgv_nhucau.ColumnHeadersHeight / 2
        Me.dgv_nhucau.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_nhucau_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        Dim rCotDauHeight As Integer = 0

        'Chỉ lặp những cột nào cần merge
        For j As Integer = 0 To dgv_nhucau.ColumnCount - 1 Step 1
            If j = 9 Then
                Dim rCotDau As Rectangle = Me.dgv_nhucau.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 2, -1, True).Width
                rCotDau.X += 1
                rCotDau.Y += 1
                rCotDau.Width = rCotDau.Width + w1 + w2 - 1
                rCotDau.Height = rCotDau.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                rCotDauHeight = rCotDau.Height
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.BackColor), rCotDau)
                e.Graphics.DrawRectangle(Pens.Silver, rCotDau)
                e.Graphics.DrawString("Trong đó", Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Font,
                                      New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.ForeColor), rCotDau, format)
            End If
            If j = 13 Then
                Dim rCotHai As Rectangle = Me.dgv_nhucau.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 1, -1, True).Width
                rCotHai.X += 1
                rCotHai.Y += 1
                rCotHai.Width = rCotHai.Width + w1 - 1
                rCotHai.Height = rCotHai.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.BackColor), rCotHai)
                e.Graphics.DrawRectangle(Pens.Silver, rCotHai)
                e.Graphics.DrawString("Trong đó", Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Font,
                                                     New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.ForeColor), rCotHai, format)
            End If
            If j = 15 Then       'Mạng lưới đơn vị
                Dim rCotBa As Rectangle = Me.dgv_nhucau.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 3, -1, True).Width
                Dim w4 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 4, -1, True).Width
                rCotBa.X += 1
                rCotBa.Y += 1
                rCotBa.Width = rCotBa.Width + w1 + w2 + w3 + w4 - 1
                rCotBa.Height = rCotBa.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                rCotDauHeight = rCotBa.Height
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.BackColor), rCotBa)
                e.Graphics.DrawRectangle(Pens.Silver, rCotBa)
                e.Graphics.DrawString("Mạng lưới đơn vị", Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Font,
                                      New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.ForeColor), rCotBa, format)
            End If
            If j = 21 Then       '
                Dim rCotBa As Rectangle = Me.dgv_nhucau.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 1, -1, True).Width
                rCotBa.X += 1
                rCotBa.Y += 1
                rCotBa.Width = rCotBa.Width + w1 - 1
                rCotBa.Height = (rCotBa.Height / _ClassNumber - 2) * 2 ' /3 la vi 3 lop
                rCotDauHeight = rCotBa.Height
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.BackColor), rCotBa)
                e.Graphics.DrawRectangle(Pens.Silver, rCotBa)
                e.Graphics.DrawString("Cán bộ được TB trúng tuyển vào đơn vị trong tháng đang tham gia lớp đào tạo", Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Font,
                                      New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.ForeColor), rCotBa, format)
            End If
            If j = 23 Then       '
                Dim rCotBa As Rectangle = Me.dgv_nhucau.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_nhucau.GetCellDisplayRectangle(j + 1, -1, True).Width
                rCotBa.X += 1
                rCotBa.Y += 1
                rCotBa.Width = rCotBa.Width + w1 - 1
                rCotBa.Height = (rCotBa.Height / _ClassNumber - 2) * 2 ' /3 la vi 3 lop
                rCotDauHeight = rCotBa.Height
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.BackColor), rCotBa)
                e.Graphics.DrawRectangle(Pens.Silver, rCotBa)
                e.Graphics.DrawString("Nhu cầu lao động cần bổ sung", Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Font,
                                      New SolidBrush(Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.ForeColor), rCotBa, format)
            End If
        Next
    End Sub

#End Region

#Region "---> Functions Merget DataGridView <---"
    Private Sub SetLayout_MangLuoi()
        AddHandler Me.dgv_mangluoi.CellPainting, AddressOf dgv_mangluoi_CellPainting
        AddHandler Me.dgv_mangluoi.Paint, AddressOf dgv_mangluoi_Paint
        AddHandler Me.dgv_mangluoi.Scroll, AddressOf dgv_mangluoi_Scroll
        AddHandler Me.dgv_mangluoi.ColumnWidthChanged, AddressOf dgv_mangluoi_ColumnWidthChanged
    End Sub

    Private Sub dgv_mangluoi_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.DataGridViewColumnEventArgs)
        Dim rtHeader As System.Drawing.Rectangle = Me.dgv_mangluoi.DisplayRectangle
        rtHeader.Height = Me.dgv_mangluoi.ColumnHeadersHeight / 2
        Me.dgv_mangluoi.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_mangluoi_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs)
        Dim rtHeader As System.Drawing.Rectangle = Me.dgv_mangluoi.DisplayRectangle
        rtHeader.Height = Me.dgv_mangluoi.ColumnHeadersHeight / 2
        Me.dgv_mangluoi.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_mangluoi_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            Dim r2 As System.Drawing.Rectangle = e.CellBounds
            r2.Y += e.CellBounds.Height / 2
            r2.Height = e.CellBounds.Height / 2
            e.PaintBackground(r2, True)
            e.PaintContent(r2)
            e.Handled = True
        End If
    End Sub

    Private Sub dgv_mangluoi_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Merge("Dư nợ (đơn vị: tỷ đồng)", 11, 14, e)
    End Sub

    Dim r1 As System.Drawing.Rectangle, r2 As System.Drawing.Rectangle, r3 As System.Drawing.Rectangle
    Private Sub Merge(ByVal _title As String, ByVal _start As Integer, ByVal _end As Integer, ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center

        r1 = dgv_mangluoi.GetCellDisplayRectangle(_start, -1, True)
        r2 = dgv_mangluoi.GetCellDisplayRectangle(_end - 1, -1, True)
        r3 = r1

        r3.Height = r1.Height / 2
        r3.Y = r1.Y + 2         'Top
        r3.X = r1.X + 2         'Left
        r3.Width = r2.Width + r2.X - r1.X
        r2 = dgv_mangluoi.GetCellDisplayRectangle(_end, -1, True)
        r3.Width = r3.Width + r2.Width - 4

        e.Graphics.FillRectangle(New SolidBrush(dgv_mangluoi.ColumnHeadersDefaultCellStyle.BackColor), r3.X, r3.Y, r3.Width, r3.Height)
        e.Graphics.DrawLine(New Pen(dgv_mangluoi.GridColor, 1), r3.X, r3.Bottom, r3.X + r3.Width, r3.Bottom)
        e.Graphics.DrawString(_title, dgv_mangluoi.ColumnHeadersDefaultCellStyle.Font, New SolidBrush(dgv_mangluoi.ColumnHeadersDefaultCellStyle.ForeColor), r3, format)
    End Sub
#End Region

#Region "---> Events Main <---"
    Private Sub KhLd_MangLuoiForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Kế hoạch lao động
        IdKHML_Sele = ""
        _KHLD_MangLuoiBLL.Create_Frame_GridView(dgv_khld, 1)
        Dim _Loai_DL As Byte = IIf(_FlagTWCN = 1, 2, 3)
        cb_khld_thoidiem_xemdl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(_Loai_DL, "---Chọn thời điểm---", _LoginPos)

        cb_khld_thoidiem_xemdl.DisplayMember = "Display"
        cb_khld_thoidiem_xemdl.ValueMember = "Value"
        If Not (cb_khld_thoidiem_xemdl Is Nothing) Then
            If cb_khld_thoidiem_xemdl.Items.Count > 1 Then
                cb_khld_thoidiem_xemdl.SelectedIndex = 1
            End If
        End If

        ARL_KhLd_DonVi.Clear()
        cb_khld_donvi_tb.Items.Clear()
        ARL_KhLd_DonVi = _Globals.Bind_ComBoBox(cb_khld_donvi_tb, ChiNhanhBLL.GetQuerySQL_Branch(_LoginPos, 0, IIf(_FlagTWCN = 1, 0, 1)), "---Chọn đơn vị---")
        If Not (cb_khld_donvi_tb Is Nothing) Then
            If cb_khld_donvi_tb.Items.Count > 1 Then
                cb_khld_donvi_tb.SelectedIndex = 0
            End If
        End If

        edt_khld_sodaihan.Text = "0"
        edt_khld_songanhan.Text = "0"
        dtpk_khld_ngaytb.Text = DateTime.Now.ToShortDateString()
        edt_khld_socongvan.Text = ""
        edt_khld_ghichu.Text = ""

        'Mạng lưới Đơn vị
        _KHLD_MangLuoiBLL.Create_Frame_GridView(dgv_mangluoi, 0)
        SetLayout_MangLuoi()
        cb_ml_thoidiem_dl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(1, "---Chọn thời điểm---", _LoginPos)
        cb_ml_thoidiem_dl.DisplayMember = "Display"
        cb_ml_thoidiem_dl.ValueMember = "Value"
        If Not (cb_ml_thoidiem_dl Is Nothing) Then
            If cb_ml_thoidiem_dl.Items.Count > 1 Then
                cb_ml_thoidiem_dl.SelectedIndex = 1
            End If
        End If

        'Nhu cầu lao động cần bổ sung và Số cán bộ được thông báo trúng tuyển vào chi nhánh trong tháng đang tham gia lớp đào tạo
        _KHLD_MangLuoiBLL.Create_Frame_GridView(dgv_nhucau, 2)
        Me.dgv_nhucau.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        Me.dgv_nhucau.ColumnHeadersHeight = Me.dgv_nhucau.ColumnHeadersHeight * _ClassNumber
        Me.dgv_nhucau.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        AddHandler Me.dgv_nhucau.CellPainting, AddressOf dgv_nhucau_CellPainting
        AddHandler Me.dgv_nhucau.Paint, AddressOf dgv_nhucau_Paint
        AddHandler Me.dgv_nhucau.Scroll, AddressOf dgv_nhucau_Scroll
        AddHandler Me.dgv_nhucau.ColumnWidthChanged, AddressOf dgv_nhucau_ColumnWidthChanged


        cb_nhucau_thoidiem_xemdl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(4, "---Chọn thời điểm---", _LoginPos)

        cb_nhucau_thoidiem_xemdl.DisplayMember = "Display"
        cb_nhucau_thoidiem_xemdl.ValueMember = "Value"
        If Not (cb_nhucau_thoidiem_xemdl Is Nothing) Then
            If cb_nhucau_thoidiem_xemdl.Items.Count > 1 Then
                cb_nhucau_thoidiem_xemdl.SelectedIndex = 1
            End If
        End If


        Dim _DateReport As DateTime = Globals.GetDateTime_ForServerDB
        dtpk_nhucau_ngaybc.Value = DateTimeUtil.GetStartOfMonth(_DateReport.Month, _DateReport.Year)
        ARL_NhuCau_DonVi.Clear()
        cb_nhucau_donvi.Items.Clear()
        ARL_NhuCau_DonVi = _Globals.Bind_ComBoBox(cb_nhucau_donvi, ChiNhanhBLL.GetQuerySQL_Branch(_LoginPos, 2, IIf(_FlagTWCN = 1, 0, 1)), "---Chọn đơn vị---")
        If Not (cb_nhucau_donvi Is Nothing) Then
            If cb_nhucau_donvi.Items.Count > 1 Then
                cb_nhucau_donvi.SelectedIndex = 1
            End If
        End If

        tctrl_main_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub tctrl_main_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tctrl_main.SelectedIndexChanged
        lbl_tongsobg.Text = ""
        Dim iCountRows As Integer = 0
        _FlagAdd = False
        dtpk_khld_ngaytb.Enabled = False
        cb_khld_donvi_tb.Enabled = False
        edt_khld_sodaihan.Text = "0"
        edt_khld_songanhan.Text = "0"
        dtpk_khld_ngaytb.Text = DateTime.Now.ToShortDateString()
        edt_khld_socongvan.Text = ""
        edt_khld_ghichu.Text = ""

        Select Case tctrl_main.SelectedTab.Name
            Case "tp_khld"
                btn_themmoi.Visible = True
                lbl_div_button1.Visible = True
                btn_luulai.Visible = True
                lbl_div_button2.Visible = True
                btn_huybo.Visible = True
                lbl_div_button3.Visible = True
                btn_xoabo.Visible = True
                lbl_div_button4.Visible = True
                cb_khld_thoidiem_xemdl.Enabled = True

                cb_khld_thoidiem_xemdl_SelectedIndexChanged(sender, e)
            Case "tp_mangluoi"
                btn_themmoi.Visible = False
                lbl_div_button1.Visible = False
                btn_luulai.Visible = False
                lbl_div_button2.Visible = False
                btn_huybo.Visible = False
                lbl_div_button3.Visible = False
                btn_xoabo.Visible = False
                lbl_div_button4.Visible = False

                cb_ml_thoidiem_dl_SelectedIndexChanged(sender, e)
            Case "tp_nhucau"
                btn_themmoi.Visible = True
                lbl_div_button1.Visible = True
                btn_luulai.Visible = True
                lbl_div_button2.Visible = True
                btn_huybo.Visible = True
                lbl_div_button3.Visible = True
                btn_xoabo.Visible = True
                lbl_div_button4.Visible = True
                dtpk_nhucau_ngaybc.Enabled = True
                cb_nhucau_thoidiem_xemdl_SelectedIndexChanged(sender, e)
        End Select
    End Sub

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub btn_themmoi_Click(sender As Object, e As EventArgs) Handles btn_themmoi.Click
        _FlagAdd = True
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_khld"
                edt_khld_sodaihan.Text = "0"
                edt_khld_songanhan.Text = "0"
                dtpk_khld_ngaytb.Text = DateTime.Now.ToShortDateString()
                edt_khld_socongvan.Text = ""
                edt_khld_ghichu.Text = ""

                cb_khld_thoidiem_xemdl.Enabled = False
                dtpk_khld_ngaytb.Enabled = True
                cb_khld_donvi_tb.Enabled = True
            Case "tp_nhucau"
                Dim _DateReport As DateTime = Globals.GetDateTime_ForServerDB
                dtpk_nhucau_ngaybc.Value = DateTimeUtil.GetStartOfMonth(_DateReport.Month, _DateReport.Year)
                dtpk_nhucau_ngaybc.Enabled = True
                If Not (cb_nhucau_donvi Is Nothing) Then
                    If cb_nhucau_donvi.Items.Count > 1 Then
                        cb_nhucau_donvi.SelectedIndex = 1
                    End If
                End If
                edt_nhucau_nhucaubs_sl.Text = "0"
                edt_nhucau_nhucaubs_socv.Text = ""
                edt_nhucau_thbaottdangdt_sl.Text = "0"
                edt_nhucau_thbaottdangdt_socv.Text = ""
                edt_nhucau_ghichu.Text = ""
        End Select
    End Sub

    Private Sub btn_xoabo_Click(sender As Object, e As EventArgs) Handles btn_xoabo.Click
        If (dgv_khld.Rows.Count <= 0) Then Return
        Try
            If (Not (dgv_khld Is Nothing) And dgv_khld.Rows.Count > 0) Then
                If ((dgv_khld.CurrentRow.Cells("cln_STT").Value IsNot Nothing) And (dgv_khld.CurrentRow.Cells("cln_STT").Value.ToString() <> "")) Then
                    Dim sIdKHML As String = dgv_khld.CurrentRow.Cells("cln_IdKHML").Value.ToString().Trim()
                    If (_KHLD_MangLuoiBLL.Delete_KhLd_MangLuoi(sIdKHML, gUsername)) = True Then
                        MessageBox.Show("Bạn đã xoá thành công Thông báo KHLĐ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                        Dim _Loai_DL As Byte = IIf(_FlagTWCN = 1, 2, 3)
                        cb_khld_thoidiem_xemdl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(_Loai_DL, "---Chọn thời điểm---", _LoginPos)

                        cb_khld_thoidiem_xemdl.DisplayMember = "Display"
                        cb_khld_thoidiem_xemdl.ValueMember = "Value"
                        If Not (cb_khld_thoidiem_xemdl Is Nothing) Then
                            If cb_khld_thoidiem_xemdl.Items.Count > 1 Then
                                cb_khld_thoidiem_xemdl.SelectedIndex = 1
                            End If
                        End If

                        ARL_KhLd_DonVi.Clear()
                        cb_khld_donvi_tb.Items.Clear()
                        ARL_KhLd_DonVi = _Globals.Bind_ComBoBox(cb_khld_donvi_tb, ChiNhanhBLL.GetQuerySQL_Branch(_LoginPos, 0, IIf(_FlagTWCN = 1, 0, 1)), "---Chọn đơn vị---")
                        tctrl_main_SelectedIndexChanged(sender, e)
                    End If

                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Xoá thông báo KHLĐ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_luulai_Click(sender As Object, e As EventArgs) Handles btn_luulai.Click
        Dim _KHLD As KHLD_MangLuoi.KhLd_MangLuoi = New KHLD_MangLuoi.KhLd_MangLuoi()

        Select Case tctrl_main.SelectedTab.Name
            Case "tp_khld"      ' Kế hoạch lao động
                _KHLD.ThoiDiem = CType(dtpk_khld_ngaytb.Value.ToString("yyyy-MM-dd"), DateTime)
                _KHLD.Loai_DL = IIf(_FlagTWCN = 1, 2, 3)
                _KHLD.IdDonVi = CType(IIf(ARL_KhLd_DonVi.Count > 0, ARL_KhLd_DonVi(cb_khld_donvi_tb.SelectedIndex), "0"), Integer)
                _KHLD.DonVi_Cd = SoftSqlHelper.GetString(String.Format("Select Ma_So From ChiNhanh Where Id={0}", _KHLD.IdDonVi), "")
                If String.IsNullOrEmpty(_KHLD.DonVi_Cd) Then
                    MessageBox.Show("Đơn vị '" & cb_khld_donvi_tb.SelectedItem.ToString().Trim() & "' cần nhập KHLĐ chưa có mã POS. Vui lòng liên hệ quản trị!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                Else
                    _KHLD.IdPhongBan = 0
                    _KHLD.SoLD_DaiHan = CType(edt_khld_sodaihan.Text.Trim().Replace(",", "").Replace(" ", ""), Integer)
                    _KHLD.SoLD_NganHan = CType(edt_khld_songanhan.Text.Trim().Replace(",", "").Replace(" ", ""), Integer)
                    _KHLD.So_XaPhuong = 0
                    _KHLD.So_DiemGD = 0
                    _KHLD.So_ToTKVV = 0
                    _KHLD.SoKH_DN = 0
                    _KHLD.TongDN = 0
                    _KHLD.DuNo_TH = 0
                    _KHLD.DuNo_QH = 0
                    _KHLD.DuNo_KH = 0
                    _KHLD.SoCV_ThBao = edt_khld_socongvan.Text.Trim()
                    _KHLD.NhuCauBS_SL = 0
                    _KHLD.NhuCauBS_SoCV = ""
                    _KHLD.ThBaoTTDangDT_SL = 0
                    _KHLD.ThBaoTTDangDT_SoCV = ""
                    _KHLD.GhiChu = edt_khld_ghichu.Text.Trim()
                    _KHLD.CreatedBy = gUsername
                    _KHLD.ModifiedBy = gUsername
                    _KHLD.IdKHML = IdKHML_Sele
                    If (IsValid(_KHLD.DonVi_Cd, _KHLD.ThoiDiem, _KHLD.Loai_DL, _KHLD.SoLD_DaiHan, _KHLD.SoLD_NganHan, _KHLD.NhuCauBS_SL, _KHLD.ThBaoTTDangDT_SL, _KHLD.IdKHML, 2)) Then
                        If _FlagAdd = True Then     'Trường hợp thêm mới DL
                            _KHLD.IdKHML = ""
                            _KHLD.IdKHML = _KHLD_MangLuoiBLL.Insert_Update_KhLd_MangLuoi(_KHLD)
                        Else                        'Trường hợp sửa đổi DL
                            _KHLD.IdKHML = _KHLD_MangLuoiBLL.Insert_Update_KhLd_MangLuoi(_KHLD)
                        End If
                        MessageBox.Show("Câp nhật kế hoạch lao động của đơn vị " & cb_khld_donvi_tb.SelectedItem.ToString().Trim() & " thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                        Dim _Loai_DL As Byte = IIf(_FlagTWCN = 1, 2, 3)
                        cb_khld_thoidiem_xemdl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(_Loai_DL, "---Chọn thời điểm---", _LoginPos)

                        cb_khld_thoidiem_xemdl.DisplayMember = "Display"
                        cb_khld_thoidiem_xemdl.ValueMember = "Value"
                        If Not (cb_khld_thoidiem_xemdl Is Nothing) Then
                            If cb_khld_thoidiem_xemdl.Items.Count > 1 Then
                                cb_khld_thoidiem_xemdl.SelectedIndex = 1
                            End If
                        End If

                        ARL_KhLd_DonVi.Clear()
                        cb_khld_donvi_tb.Items.Clear()
                        ARL_KhLd_DonVi = _Globals.Bind_ComBoBox(cb_khld_donvi_tb, ChiNhanhBLL.GetQuerySQL_Branch(_LoginPos, 0, IIf(_FlagTWCN = 1, 0, 1)), "---Chọn đơn vị---")
                        tctrl_main_SelectedIndexChanged(sender, e)
                    End If
                End If
            Case "tp_nhucau"    ' Nhu cầu đào tạo
                _KHLD.ThoiDiem = DateTimeUtil.GetDateEndOfMonth(dtpk_nhucau_ngaybc.Value.Month, dtpk_nhucau_ngaybc.Value.Year)
                _KHLD.Loai_DL = 4
                _KHLD.IdDonVi = CType(IIf(ARL_NhuCau_DonVi.Count > 0, ARL_NhuCau_DonVi(cb_nhucau_donvi.SelectedIndex), "0"), Integer)
                _KHLD.DonVi_Cd = SoftSqlHelper.GetString(String.Format("Select Ma_So From ChiNhanh Where Id={0}", _KHLD.IdDonVi), "")
                If String.IsNullOrEmpty(_KHLD.DonVi_Cd) Then
                    MessageBox.Show("Đơn vị '" & cb_nhucau_donvi.SelectedItem.ToString().Trim() & "' cần nhập KHLĐ chưa có mã POS. Vui lòng liên hệ quản trị!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                Else
                    _KHLD.IdPhongBan = 0
                    _KHLD.SoLD_DaiHan = 0
                    _KHLD.SoLD_NganHan = 0
                    _KHLD.So_XaPhuong = 0
                    _KHLD.So_DiemGD = 0
                    _KHLD.So_ToTKVV = 0
                    _KHLD.SoKH_DN = 0
                    _KHLD.TongDN = 0
                    _KHLD.DuNo_TH = 0
                    _KHLD.DuNo_QH = 0
                    _KHLD.DuNo_KH = 0
                    _KHLD.SoCV_ThBao = ""
                    If String.IsNullOrEmpty(edt_nhucau_nhucaubs_sl.Text.Trim().Replace(",", "").Replace(" ", "")) Then
                        _KHLD.NhuCauBS_SL = 0
                    Else : _KHLD.NhuCauBS_SL = CType(edt_nhucau_nhucaubs_sl.Text.Trim().Replace(",", "").Replace(" ", ""), Integer)
                    End If
                    _KHLD.NhuCauBS_SoCV = edt_nhucau_nhucaubs_socv.Text.Trim()
                    If String.IsNullOrEmpty(edt_nhucau_thbaottdangdt_sl.Text.Trim().Replace(",", "").Replace(" ", "")) Then
                        _KHLD.ThBaoTTDangDT_SL = 0
                    Else : _KHLD.ThBaoTTDangDT_SL = CType(edt_nhucau_thbaottdangdt_sl.Text.Trim().Replace(",", "").Replace(" ", ""), Integer)
                    End If
                    _KHLD.ThBaoTTDangDT_SoCV = edt_nhucau_thbaottdangdt_socv.Text.Trim()
                    _KHLD.GhiChu = edt_nhucau_ghichu.Text.Trim()
                    _KHLD.CreatedBy = gUsername
                    _KHLD.ModifiedBy = gUsername
                    _KHLD.IdKHML = IdKHML_Sele
                    If (IsValid(_KHLD.DonVi_Cd, _KHLD.ThoiDiem, _KHLD.Loai_DL, _KHLD.SoLD_DaiHan, _KHLD.SoLD_NganHan, _KHLD.NhuCauBS_SL, _KHLD.ThBaoTTDangDT_SL, _KHLD.IdKHML, 4)) Then
                        If _FlagAdd = True Then     'Trường hợp thêm mới DL
                            _KHLD.IdKHML = ""
                            _KHLD.IdKHML = _KHLD_MangLuoiBLL.Insert_Update_KhLd_MangLuoi(_KHLD)
                        Else                        'Trường hợp sửa đổi DL
                            _KHLD.IdKHML = _KHLD_MangLuoiBLL.Insert_Update_KhLd_MangLuoi(_KHLD)
                        End If
                        MessageBox.Show("Câp nhật Nhu cầu lđ bổ sung - Thông báo trúng tuyển vào trong tháng đang tham gia đào tạo của đơn vị " & cb_nhucau_donvi.SelectedItem.ToString().Trim() & " thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Dim _Loai_DL As Byte = 4
                        cb_nhucau_thoidiem_xemdl.DataSource = _KHLD_MangLuoiBLL.GetListThoiDiem(_Loai_DL, "---Chọn thời điểm---", _LoginPos)
                        cb_nhucau_thoidiem_xemdl.DisplayMember = "Display"
                        cb_nhucau_thoidiem_xemdl.ValueMember = "Value"
                        If Not (cb_nhucau_thoidiem_xemdl Is Nothing) Then
                            If cb_nhucau_thoidiem_xemdl.Items.Count > 1 Then
                                cb_nhucau_thoidiem_xemdl.SelectedIndex = 1
                            End If
                        End If
                        tctrl_main_SelectedIndexChanged(sender, e)
                    End If
                End If
        End Select


    End Sub

    Private Sub cb_khld_thoidiem_xemdl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_khld_thoidiem_xemdl.SelectedIndexChanged
        lbl_tongsobg.Text = ""
        dgv_khld.Rows.Clear()
        IdKHML_Sele = ""
        Dim iCountRows As Integer = 0
        If (Not (cb_khld_thoidiem_xemdl Is Nothing) And (cb_khld_thoidiem_xemdl.SelectedIndex > 0) And (cb_khld_thoidiem_xemdl.Items.Count <> 0)) Then
            Dim _ThoiDiemVN As String = cb_khld_thoidiem_xemdl.Text.ToString()
            If Not String.IsNullOrEmpty(_ThoiDiemVN) And _ThoiDiemVN.Length = 10 Then
                Dim arrTemp() As String = _ThoiDiemVN.Split("-")
                Dim _ThoiDiemEN As String = arrTemp(2).ToString() + "-" + arrTemp(1).ToString() + "-" + arrTemp(0).ToString()
                Dim _DonVi_Cd As String = IIf(_FlagTWCN = 1, "", _LoginPos)
                Dim _Loai_DL As Byte = IIf(_FlagTWCN = 1, 2, 3)
                Dim _KieuIn As Byte = 0
                Using db As System.Data.DataTable = _KHLD_MangLuoiBLL.GetKhLd_MangLuoi_Search(_FlagTWCN, _ThoiDiemEN, 0, "", _Loai_DL, 0, _DonVi_Cd)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_khld.Rows.Add()
                                dgv_khld.Rows(i).Cells("cln_IdKHML").Value = db.Rows(i)("IdKHML").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_Id").Value = db.Rows(i)("Id").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_KieuIn").Value = db.Rows(i)("KieuIn").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_IdDonVi").Value = db.Rows(i)("IdDonVi").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_ChiNhanh_Id").Value = db.Rows(i)("ChiNhanh_Id").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_IdPhongBan").Value = db.Rows(i)("IdPhongBan").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_Loai_DL").Value = db.Rows(i)("Loai_DL").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_DonVi_Cd").Value = db.Rows(i)("DonVi_Cd").ToString().Trim()

                                If (db.Rows(i)("ThoiDiem").ToString().Trim() <> "" And Format(CType(db.Rows(i)("ThoiDiem"), DateTime), "dd-MM-yyyy") <> "01-01-1900") Then
                                    dgv_khld.Rows(i).Cells("cln_ThoiDiem").Value = Format(CType(db.Rows(i)("ThoiDiem"), DateTime), "dd-MM-yyyy")
                                Else
                                    dgv_khld.Rows(i).Cells("cln_ThoiDiem").Value = ""
                                End If
                                dgv_khld.Rows(i).Cells("cln_STT").Value = IIf(db.Rows(i)("STT").ToString().Trim() = "0", "", db.Rows(i)("STT").ToString().Trim())

                                dgv_khld.Rows(i).Cells("cln_Name").Value = db.Rows(i)("Name").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_SoLD_DaiHan").Value = IIf(db.Rows(i)("SoLD_DaiHan").ToString() <> "", Double.Parse(db.Rows(i)("SoLD_DaiHan").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")
                                dgv_khld.Rows(i).Cells("cln_SoLD_NganHan").Value = IIf(db.Rows(i)("SoLD_NganHan").ToString() <> "", Double.Parse(db.Rows(i)("SoLD_NganHan").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")
                                dgv_khld.Rows(i).Cells("cln_SoCV_ThBao").Value = db.Rows(i)("SoCV_ThBao").ToString().Trim()
                                dgv_khld.Rows(i).Cells("cln_GhiChu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                                'Chỉnh style row trên GridView hiển thị
                                _KieuIn = CType(db.Rows(i)("KieuIn"), Byte)
                                SetStyleRowGrid(i, dgv_khld, _KieuIn)
                                iCountRows = iCountRows + 1
                            Next
                        End If
                    End If
                    dgv_khld_CellClick(Nothing, Nothing)
                End Using
            End If
        End If
        lbl_tongsobg.Text = " Số bản ghi: " & iCountRows.ToString()
    End Sub

    Private Sub cb_ml_thoidiem_dl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ml_thoidiem_dl.SelectedIndexChanged
        lbl_tongsobg.Text = ""
        dgv_mangluoi.Rows.Clear()
        Dim _KieuIn As Byte = 0
        Dim iCountRows As Integer = 0

        If (Not (cb_ml_thoidiem_dl Is Nothing) And (cb_ml_thoidiem_dl.SelectedIndex > 0) And (cb_ml_thoidiem_dl.Items.Count <> 0)) Then
            Dim _ThoiDiemVN As String = cb_ml_thoidiem_dl.Text.ToString()
            If Not String.IsNullOrEmpty(_ThoiDiemVN) And _ThoiDiemVN.Length = 10 Then
                Dim arrTemp() As String = _ThoiDiemVN.Split("-")
                Dim _ThoiDiemEN As String = arrTemp(2).ToString() + "-" + arrTemp(1).ToString() + "-" + arrTemp(0).ToString()
                Dim _DonVi_Cd As String = IIf(_FlagTWCN = 1, "", IIf(_FlagTWCN = 2 And _LoginPos <> "000101" And _LoginPos <> "000196" And _LoginPos <> "000197", _LoginPos.Substring(0, 4), _LoginPos))
                Using db As System.Data.DataTable = _KHLD_MangLuoiBLL.GetKhLd_MangLuoi_Search(_FlagTWCN, _ThoiDiemEN, 0, "", 1, 0, _DonVi_Cd)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_mangluoi.Rows.Add()
                                dgv_mangluoi.Rows(i).Cells("cln_IdKHML").Value = db.Rows(i)("IdKHML").ToString().Trim()
                                dgv_mangluoi.Rows(i).Cells("cln_Id").Value = db.Rows(i)("Id").ToString().Trim()
                                dgv_mangluoi.Rows(i).Cells("cln_STT").Value = IIf(db.Rows(i)("STT").ToString().Trim() = "0", "", db.Rows(i)("STT").ToString().Trim())
                                dgv_mangluoi.Rows(i).Cells("cln_KieuIn").Value = db.Rows(i)("KieuIn").ToString().Trim()

                                If (db.Rows(i)("ThoiDiem").ToString().Trim() <> "") Then
                                    dgv_mangluoi.Rows(i).Cells("cln_ThoiDiem").Value = Format(CType(db.Rows(i)("ThoiDiem"), DateTime), "dd-MM-yyyy")
                                Else
                                    dgv_mangluoi.Rows(i).Cells("cln_ThoiDiem").Value = ""
                                End If
                                dgv_mangluoi.Rows(i).Cells("cln_DonVi").Value = db.Rows(i)("DonVi_HT").ToString().Trim()
                                dgv_mangluoi.Rows(i).Cells("cln_SoXaPhuong").Value = IIf(db.Rows(i)("So_XaPhuong").ToString() <> "", Double.Parse(db.Rows(i)("So_XaPhuong").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_SoDiemGD").Value = IIf(db.Rows(i)("So_DiemGD").ToString() <> "", Double.Parse(db.Rows(i)("So_DiemGD").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_SoToTKVV").Value = IIf(db.Rows(i)("So_ToTKVV").ToString() <> "", Double.Parse(db.Rows(i)("So_ToTKVV").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_SoKH_DN").Value = IIf(db.Rows(i)("SoKH_DN").ToString() <> "", Double.Parse(db.Rows(i)("SoKH_DN").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D0), "0")

                                dgv_mangluoi.Rows(i).Cells("cln_TongDN").Value = IIf(db.Rows(i)("TongDN").ToString() <> "", Double.Parse(db.Rows(i)("TongDN").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D2), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_DuNo_TH").Value = IIf(db.Rows(i)("DuNo_TH").ToString() <> "", Double.Parse(db.Rows(i)("DuNo_TH").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D2), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_DuNo_QH").Value = IIf(db.Rows(i)("DuNo_QH").ToString() <> "", Double.Parse(db.Rows(i)("DuNo_QH").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D2), "0")
                                dgv_mangluoi.Rows(i).Cells("cln_DuNo_KH").Value = IIf(db.Rows(i)("DuNo_KH").ToString() <> "", Double.Parse(db.Rows(i)("DuNo_KH").ToString(), Globals.cultureNum).ToString("N", vNFInfo_D2), "0")

                                dgv_mangluoi.Rows(i).Cells("cln_GhiChu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                                'Chỉnh style row trên GridView hiển thị
                                _KieuIn = CType(db.Rows(i)("KieuIn"), Byte)
                                SetStyleRowGrid(i, dgv_mangluoi, _KieuIn)
                                iCountRows = iCountRows + 1
                            Next
                        End If
                    End If
                End Using
            End If
        End If
        lbl_tongsobg.Text = " Số bản ghi: " & iCountRows.ToString()
    End Sub

    Private Sub cb_nhucau_thoidiem_xemdl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_nhucau_thoidiem_xemdl.SelectedIndexChanged
        lbl_tongsobg.Text = ""
        dgv_nhucau.Rows.Clear()
        IdKHML_Sele = ""
        Dim iCountRows As Integer = 0
        If (Not (cb_nhucau_thoidiem_xemdl Is Nothing) And (cb_nhucau_thoidiem_xemdl.Items.Count <> 0)) Then
            Dim _ThoiDiemVN As String = "", _ThoiDiemEN As String = ""
            If cb_nhucau_thoidiem_xemdl.SelectedIndex > 0 Then
                _ThoiDiemVN = cb_nhucau_thoidiem_xemdl.Text.ToString()
            End If
            If _ThoiDiemVN <> "" Then
                Dim arrTemp() As String = _ThoiDiemVN.Split("-")
                _ThoiDiemEN = arrTemp(2).ToString() + "-" + arrTemp(1).ToString() + "-" + arrTemp(0).ToString()
            End If

            Dim _DonVi_Cd As String = IIf(_FlagTWCN = 1, "", _LoginPos)
            Dim _KieuIn As Byte = 0
            Using db As System.Data.DataTable = _KHLD_MangLuoiBLL.GetKhLd_MangLuoi_Search(_FlagTWCN, _ThoiDiemEN, 0, "", 4, 0, _DonVi_Cd)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        For i As Integer = 0 To db.Rows.Count - 1
                            dgv_nhucau.Rows.Add()
                            dgv_nhucau.Rows(i).Cells("cln_IdKHML").Value = db.Rows(i)("IdKHML").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_Id").Value = db.Rows(i)("Id").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_KieuIn").Value = db.Rows(i)("KieuIn").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_STT").Value = db.Rows(i)("STT").ToString().Trim()
                            If (db.Rows(i)("ThoiDiem").ToString().Trim() <> "" And Format(CType(db.Rows(i)("ThoiDiem"), DateTime), "dd-MM-yyyy") <> "01-01-1900") Then
                                dgv_nhucau.Rows(i).Cells("cln_ThoiDiem").Value = Format(CType(db.Rows(i)("ThoiDiem"), DateTime), "dd-MM-yyyy")
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_ThoiDiem").Value = ""
                            End If
                            dgv_nhucau.Rows(i).Cells("cln_DonVi_HT").Value = db.Rows(i)("DonVi_HT").ToString().Trim()
                            If IsNothing(db.Rows(i)("SoLD_ThgTruoc")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgTruoc").ToString()) Or db.Rows(i)("SoLD_ThgTruoc").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgTruoc").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgTruoc").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgTruoc").Value = Double.Parse(db.Rows(i)("SoLD_ThgTruoc").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_DaiHan")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_DaiHan").ToString()) Or db.Rows(i)("SoLD_ThgBC_DaiHan").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_DaiHan").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_DaiHan").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_DaiHan").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_DaiHan").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_DangLV")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_DangLV").ToString()) Or db.Rows(i)("SoLD_ThgBC_DangLV").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_DangLV").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_DangLV").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_DangLV").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_DangLV").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_Nghi_HuongBHXH")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_Nghi_HuongBHXH").ToString()) Or db.Rows(i)("SoLD_ThgBC_Nghi_HuongBHXH").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_Nghi_HuongBHXH").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_Nghi_HuongBHXH").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_Nghi_HuongBHXH").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_Nghi_HuongBHXH").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_Nghi_KhongBHXH")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_Nghi_KhongBHXH").ToString()) Or db.Rows(i)("SoLD_ThgBC_Nghi_KhongBHXH").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_Nghi_KhongBHXH").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_Nghi_KhongBHXH").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_Nghi_KhongBHXH").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_Nghi_KhongBHXH").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_NganHan")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_NganHan").ToString()) Or db.Rows(i)("SoLD_ThgBC_NganHan").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_NganHan").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_NganHan").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_NganHan_DB")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_NganHan_DB").ToString()) Or db.Rows(i)("SoLD_ThgBC_NganHan_DB").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_NganHan_DB").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan_DB").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan_DB").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_NganHan_DB").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_ThgBC_NganHan_PT")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_ThgBC_NganHan_PT").ToString()) Or db.Rows(i)("SoLD_ThgBC_NganHan_PT").ToString().Trim() = "" Or db.Rows(i)("SoLD_ThgBC_NganHan_PT").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan_PT").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_ThgBC_NganHan_PT").Value = Double.Parse(db.Rows(i)("SoLD_ThgBC_NganHan_PT").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("TongDN")) Or String.IsNullOrEmpty(db.Rows(i)("TongDN").ToString()) Or db.Rows(i)("TongDN").ToString().Trim() = "" Or db.Rows(i)("TongDN").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_TongDN").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_TongDN").Value = Double.Parse(db.Rows(i)("TongDN").ToString(), Globals.cultureNum).ToString("N2", vNFInfo_D2)
                            End If
                            If IsNothing(db.Rows(i)("SoKH_DN")) Or String.IsNullOrEmpty(db.Rows(i)("SoKH_DN").ToString()) Or db.Rows(i)("SoKH_DN").ToString().Trim() = "" Or db.Rows(i)("SoKH_DN").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoKH_DN").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoKH_DN").Value = Double.Parse(db.Rows(i)("SoKH_DN").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("So_XaPhuong")) Or String.IsNullOrEmpty(db.Rows(i)("So_XaPhuong").ToString()) Or db.Rows(i)("So_XaPhuong").ToString().Trim() = "" Or db.Rows(i)("So_XaPhuong").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoXaPhuong").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoXaPhuong").Value = Double.Parse(db.Rows(i)("So_XaPhuong").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("So_DiemGD")) Or String.IsNullOrEmpty(db.Rows(i)("So_DiemGD").ToString()) Or db.Rows(i)("So_DiemGD").ToString().Trim() = "" Or db.Rows(i)("So_DiemGD").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoDiemGD").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoDiemGD").Value = Double.Parse(db.Rows(i)("So_DiemGD").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("So_ToTKVV")) Or String.IsNullOrEmpty(db.Rows(i)("So_ToTKVV").ToString()) Or db.Rows(i)("So_ToTKVV").ToString().Trim() = "" Or db.Rows(i)("So_ToTKVV").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoToTKVV").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoToTKVV").Value = Double.Parse(db.Rows(i)("So_ToTKVV").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            If IsNothing(db.Rows(i)("SoLD_GiamTuNhien")) Or String.IsNullOrEmpty(db.Rows(i)("SoLD_GiamTuNhien").ToString()) Or db.Rows(i)("SoLD_GiamTuNhien").ToString().Trim() = "" Or db.Rows(i)("SoLD_GiamTuNhien").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTuNhien_KhacCN").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTuNhien_KhacCN").Value = Double.Parse(db.Rows(i)("SoLD_GiamTuNhien_KhacCN").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If

                            If IsNothing(db.Rows(i)("ThBaoTTDangDT_SL")) Or String.IsNullOrEmpty(db.Rows(i)("ThBaoTTDangDT_SL").ToString()) Or db.Rows(i)("ThBaoTTDangDT_SL").ToString().Trim() = "" Or db.Rows(i)("ThBaoTTDangDT_SL").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_ThBaoTTDangDT_SL").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_ThBaoTTDangDT_SL").Value = Double.Parse(db.Rows(i)("ThBaoTTDangDT_SL").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            dgv_nhucau.Rows(i).Cells("cln_ThBaoTTDangDT_SoCV").Value = db.Rows(i)("ThBaoTTDangDT_SoCV").ToString().Trim()
                            If IsNothing(db.Rows(i)("NhuCauBS_SL")) Or String.IsNullOrEmpty(db.Rows(i)("NhuCauBS_SL").ToString()) Or db.Rows(i)("NhuCauBS_SL").ToString().Trim() = "" Or db.Rows(i)("NhuCauBS_SL").ToString().Trim() = "0" Then
                                dgv_nhucau.Rows(i).Cells("cln_NhuCauBS_SL").Value = ""
                            Else
                                dgv_nhucau.Rows(i).Cells("cln_NhuCauBS_SL").Value = Double.Parse(db.Rows(i)("NhuCauBS_SL").ToString(), Globals.cultureNum).ToString("N0", vNFInfo_D0)
                            End If
                            dgv_nhucau.Rows(i).Cells("cln_NhuCauBS_SoCV").Value = db.Rows(i)("NhuCauBS_SoCV").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_GhiChu").Value = db.Rows(i)("GhiChu").ToString().Trim()


                            dgv_nhucau.Rows(i).Cells("cln_Loai_DL").Value = db.Rows(i)("Loai_DL").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_Loai_DL_HT").Value = db.Rows(i)("Loai_DL_HT").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_IdDonVi").Value = db.Rows(i)("IdDonVi").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_DonVi_Cd").Value = db.Rows(i)("DonVi_Cd").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_ChiNhanh_Id").Value = db.Rows(i)("ChiNhanh_Id").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_ChiNhanh_HT").Value = db.Rows(i)("ChiNhanh_HT").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_IdPhongBan").Value = db.Rows(i)("IdPhongBan").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_PhongBan_HT").Value = db.Rows(i)("PhongBan_HT").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_DuNo_TH").Value = db.Rows(i)("DuNo_TH").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_DuNo_QH").Value = db.Rows(i)("DuNo_QH").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_DuNo_KH").Value = db.Rows(i)("DuNo_KH").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_DaiHan").Value = db.Rows(i)("SoLD_DaiHan").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_NganHan").Value = db.Rows(i)("SoLD_NganHan").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoCV_ThBao").Value = db.Rows(i)("SoCV_ThBao").ToString().Trim()

                            dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTuNhien").Value = db.Rows(i)("SoLD_GiamTuNhien").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTN_01").Value = db.Rows(i)("SoLD_GiamTN_01").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTN_KhacCN_01").Value = db.Rows(i)("SoLD_GiamTN_KhacCN_01").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTN_02").Value = db.Rows(i)("SoLD_GiamTN_02").ToString().Trim()
                            dgv_nhucau.Rows(i).Cells("cln_SoLD_GiamTN_KhacCN_02").Value = db.Rows(i)("SoLD_GiamTN_KhacCN_02").ToString().Trim()
                            'Chỉnh style row trên GridView hiển thị
                            _KieuIn = CType(db.Rows(i)("KieuIn"), Byte)
                            SetStyleRowGrid(i, dgv_nhucau, _KieuIn)
                            iCountRows = iCountRows + 1
                        Next
                    End If
                End If
                dgv_nhucau_CellClick(Nothing, Nothing)
            End Using
        End If

        lbl_tongsobg.Text = " Số bản ghi: " & iCountRows.ToString()
    End Sub

    Private Sub dgv_nhucau_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_nhucau.CellClick
        If Not (cb_nhucau_donvi Is Nothing) Then
            If cb_nhucau_donvi.Items.Count > 1 Then
                cb_nhucau_donvi.SelectedIndex = 0
            End If
        End If

        Dim _DateReport As DateTime = Globals.GetDateTime_ForServerDB
        dtpk_nhucau_ngaybc.Value = DateTimeUtil.GetStartOfMonth(_DateReport.Month, _DateReport.Year)
        If Not (cb_nhucau_donvi Is Nothing) Then
            If cb_nhucau_donvi.Items.Count > 1 Then
                cb_nhucau_donvi.SelectedIndex = 1
            End If
        End If
        edt_nhucau_nhucaubs_sl.Text = "0"
        edt_nhucau_nhucaubs_socv.Text = ""
        edt_nhucau_thbaottdangdt_sl.Text = "0"
        edt_nhucau_thbaottdangdt_socv.Text = ""
        edt_nhucau_ghichu.Text = ""
        IdKHML_Sele = ""
        Dim _DonViSelect As String = ""
        Dim _KieuIn As Byte = 3
        If (Not (dgv_nhucau Is Nothing) And dgv_nhucau.Rows.Count > 0) Then
            If ((dgv_nhucau.CurrentRow.Cells("cln_STT").Value IsNot Nothing) And (dgv_nhucau.CurrentRow.Cells("cln_STT").Value.ToString() <> "")) Then
                IdKHML_Sele = dgv_nhucau.CurrentRow.Cells("cln_IdKHML").Value.ToString().Trim()
                _KieuIn = CType(dgv_nhucau.CurrentRow.Cells("cln_KieuIn").Value.ToString().Trim(), Byte)
                If dgv_nhucau.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim() <> "" And dgv_nhucau.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim() <> "01-01-1900" Then
                    Dim arrTemp() As String = dgv_nhucau.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim().Split("-")
                    dtpk_nhucau_ngaybc.Value = New DateTime(CType(arrTemp(2).ToString(), Integer), CType(arrTemp(1).ToString(), Integer), 1)
                End If
                _DonViSelect = dgv_nhucau.CurrentRow.Cells("cln_DonVi_Cd").Value.ToString().Trim()
                cb_nhucau_donvi.SelectedIndex = IIf(dgv_nhucau.CurrentRow.Cells("cln_IdDonVi").Value.ToString().Trim() <> "", CType(ARL_NhuCau_DonVi.IndexOf(dgv_nhucau.CurrentRow.Cells("cln_IdDonVi").Value.ToString().Trim()), Integer), 0)
                If dgv_nhucau.CurrentRow.Cells("cln_NhuCauBS_SL").Value.ToString().Trim() = "" Then
                    edt_nhucau_nhucaubs_sl.Text = "0"
                Else
                    edt_nhucau_nhucaubs_sl.Text = dgv_nhucau.CurrentRow.Cells("cln_NhuCauBS_SL").Value.ToString().Trim()
                End If
                edt_nhucau_nhucaubs_socv.Text = dgv_nhucau.CurrentRow.Cells("cln_NhuCauBS_SoCV").Value.ToString().Trim()
                If dgv_nhucau.CurrentRow.Cells("cln_ThBaoTTDangDT_SL").Value.ToString().Trim() = "" Then
                    edt_nhucau_thbaottdangdt_sl.Text = "0"
                Else
                    edt_nhucau_thbaottdangdt_sl.Text = dgv_nhucau.CurrentRow.Cells("cln_ThBaoTTDangDT_SL").Value.ToString().Trim()
                End If
                edt_nhucau_thbaottdangdt_socv.Text = dgv_nhucau.CurrentRow.Cells("cln_ThBaoTTDangDT_SoCV").Value.ToString().Trim()
                edt_nhucau_ghichu.Text = dgv_nhucau.CurrentRow.Cells("cln_GhiChu").Value.ToString().Trim()
            End If
        End If

        If String.IsNullOrEmpty(IdKHML_Sele) Then
            _FlagAdd = True
            dtpk_nhucau_ngaybc.Enabled = True
        Else
            _FlagAdd = False
            dtpk_nhucau_ngaybc.Enabled = False
        End If

        If (_KieuIn = 0) Or ((DONVI = "000100" Or DONVI = "000199") And (_DonViSelect <> "000100" And _DonViSelect <> "000199")) Then
            edt_nhucau_nhucaubs_sl.ReadOnly = True
            edt_nhucau_nhucaubs_socv.ReadOnly = True
            edt_nhucau_thbaottdangdt_sl.ReadOnly = True
            edt_nhucau_thbaottdangdt_socv.ReadOnly = True
            edt_nhucau_ghichu.ReadOnly = True
        Else
            edt_nhucau_nhucaubs_sl.ReadOnly = False
            edt_nhucau_nhucaubs_socv.ReadOnly = False
            edt_nhucau_thbaottdangdt_sl.ReadOnly = False
            edt_nhucau_thbaottdangdt_socv.ReadOnly = False
            edt_nhucau_ghichu.ReadOnly = False
        End If
    End Sub

    Private Sub btn_huybo_Click(sender As Object, e As EventArgs) Handles btn_huybo.Click
        tctrl_main_SelectedIndexChanged(Nothing, Nothing)
    End Sub

    Private Sub edt_khld_sodaihan_TextChanged(sender As Object, e As EventArgs) Handles edt_khld_sodaihan.TextChanged
        edt_khld_sodaihan = formatMoneyinTextbox(edt_khld_sodaihan)
    End Sub

    Private Sub edt_khld_songanhan_TextChanged(sender As Object, e As EventArgs) Handles edt_khld_songanhan.TextChanged
        edt_khld_songanhan = formatMoneyinTextbox(edt_khld_songanhan)
    End Sub

    Private Sub dgv_khld_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_khld.CellClick
        If Not (cb_khld_donvi_tb Is Nothing) Then
            If cb_khld_donvi_tb.Items.Count > 1 Then
                cb_khld_donvi_tb.SelectedIndex = 0
            End If
        End If
        edt_khld_sodaihan.Text = "0"
        edt_khld_songanhan.Text = "0"
        dtpk_khld_ngaytb.Text = DateTime.Now.ToShortDateString()
        edt_khld_socongvan.Text = ""
        edt_khld_ghichu.Text = ""
        IdKHML_Sele = ""
        Dim _KieuIn As Byte = 0
        If (Not (dgv_khld Is Nothing) And dgv_khld.Rows.Count > 0) Then
            If ((dgv_khld.CurrentRow.Cells("cln_STT").Value IsNot Nothing) And (dgv_khld.CurrentRow.Cells("cln_STT").Value.ToString() <> "")) Then
                IdKHML_Sele = dgv_khld.CurrentRow.Cells("cln_IdKHML").Value.ToString().Trim()
                _KieuIn = CType(dgv_khld.CurrentRow.Cells("cln_KieuIn").Value.ToString().Trim(), Byte)
                If dgv_khld.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim() <> "" And dgv_khld.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim() <> "01-01-1900" Then
                    Dim arrTemp() As String = dgv_khld.CurrentRow.Cells("cln_ThoiDiem").Value.ToString().Trim().Split("-")
                    dtpk_khld_ngaytb.Value = New DateTime(CType(arrTemp(2).ToString(), Integer), CType(arrTemp(1).ToString(), Integer), CType(arrTemp(0).ToString(), Integer))
                End If
                edt_khld_sodaihan.Text = dgv_khld.CurrentRow.Cells("cln_SoLD_DaiHan").Value.ToString().Trim()
                edt_khld_songanhan.Text = dgv_khld.CurrentRow.Cells("cln_SoLD_NganHan").Value.ToString().Trim()
                edt_khld_socongvan.Text = dgv_khld.CurrentRow.Cells("cln_SoCV_ThBao").Value.ToString().Trim()
                edt_khld_ghichu.Text = dgv_khld.CurrentRow.Cells("cln_GhiChu").Value.ToString().Trim()
                cb_khld_donvi_tb.SelectedIndex = IIf(dgv_khld.CurrentRow.Cells("cln_IdDonVi").Value.ToString().Trim() <> "", CType(ARL_KhLd_DonVi.IndexOf(dgv_khld.CurrentRow.Cells("cln_IdDonVi").Value.ToString().Trim()), Integer), 0)
            End If
        End If
        If String.IsNullOrEmpty(IdKHML_Sele) Then
            _FlagAdd = True
            dtpk_khld_ngaytb.Enabled = True
        Else
            _FlagAdd = False
            dtpk_khld_ngaytb.Enabled = False
        End If

        If _KieuIn = 0 Then
            edt_khld_sodaihan.ReadOnly = True
            edt_khld_songanhan.ReadOnly = True
            edt_khld_socongvan.ReadOnly = True
            edt_khld_ghichu.ReadOnly = True
        Else
            edt_khld_sodaihan.ReadOnly = False
            edt_khld_songanhan.ReadOnly = False
            edt_khld_socongvan.ReadOnly = False
            edt_khld_ghichu.ReadOnly = False
        End If
    End Sub

    Private Sub dgv_khld_KeyUp(sender As Object, e As KeyEventArgs) Handles dgv_khld.KeyUp
        dgv_khld_CellClick(sender, Nothing)
    End Sub

    Private Sub edt_nhucau_nhucaubs_sl_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_nhucau_nhucaubs_sl.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub edt_nhucau_thbaottdangdt_sl_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_nhucau_thbaottdangdt_sl.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub edt_nhucau_nhucaubs_sl_TextChanged(sender As Object, e As EventArgs) Handles edt_nhucau_nhucaubs_sl.TextChanged
        edt_nhucau_nhucaubs_sl = formatMoneyinTextbox(edt_nhucau_nhucaubs_sl)
    End Sub

    Private Sub edt_nhucau_thbaottdangdt_sl_TextChanged(sender As Object, e As EventArgs) Handles edt_nhucau_thbaottdangdt_sl.TextChanged
        edt_nhucau_thbaottdangdt_sl = formatMoneyinTextbox(edt_nhucau_thbaottdangdt_sl)
    End Sub
#End Region

End Class