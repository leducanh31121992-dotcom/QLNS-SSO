Public Class frmChonDS_CB

    Private _listCB As List(Of String)
    Private _listCBs As String
    Private _loaiCB As Int16
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Public Property listCB() As List(Of String)
        Get
            Return _listCB
        End Get
        Set(ByVal value As List(Of String))
            _listCB = value
        End Set
    End Property

    Public Property LoaiCB() As Int16
        Get
            Return _loaiCB
        End Get
        Set(ByVal value As Int16)
            _loaiCB = value
        End Set
    End Property

    Public Property listCBs() As String
        Get
            Return _listCBs
        End Get
        Set(ByVal value As String)
            _listCBs = value
        End Set
    End Property

    Private Sub frmChonDS_CB_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmChonDS_CB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmChonDS_CB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Select Case LoaiCB
            Case "1"
                bindTreeviewFull(treeCocau, False, True, False, False, False, False, True)
            Case "2"
                bindTreeviewFull(treeCocau, False, False, True, False, False, False, True)
            Case "3"
                bindTreeviewFull(treeCocau, False, False, False, True, False)
            Case "4"
                bindTreeviewFull(treeCocau, False, False, False, False, True)
            Case Else
                bindTreeviewFull(treeCocau, True, False, False, False, False)
        End Select
        treeCocau.ExpandAll()
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub bntChoiseCB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntChoiseCB.Click
        Close()
    End Sub

    Private Sub treeCocau_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeCocau.NodeMouseClick
        If e.Node.Checked Then
            CheckSelected(e.Node, True)
        Else
            ParentUnCheckSelected(e.Node)
            CheckSelected(e.Node, False)
        End If
        listCB = getList(treeCocau.Nodes)
        'listCBs = getListCB(treeCocau.Nodes)
    End Sub

    'Private Sub rdCB_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdCB.CheckedChanged
    '    If rdCB.Checked Then
    '        LoaiCB = 0
    '        bindTreeviewFull(treeCocau, True, False, False, False, False)
    '    End If
    'End Sub

    Private Sub rdCC_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdCC.CheckedChanged
        If rdCC.Checked Then
            LoaiCB = 3
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, False, False, False, True, False)
            treeCocau.ExpandAll()
        End If
    End Sub

    Private Sub rdNH_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdNH.CheckedChanged
        If rdNH.Checked Then
            LoaiCB = 1
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, False, True, False, False, False, False, True)
            treeCocau.ExpandAll()
        End If
    End Sub

    Private Sub rdTS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdTS.CheckedChanged
        If rdTS.Checked Then
            LoaiCB = 2
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, False, False, True, False, False, False, True)
            treeCocau.ExpandAll()
        End If
    End Sub

    Private Sub rdVH_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdVH.CheckedChanged
        If rdVH.Checked Then
            LoaiCB = 4
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, False, False, False, False, True)
            treeCocau.ExpandAll()
        End If
    End Sub

    Private Sub rdCB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdCB.Click
        If rdCB.Checked Then
            LoaiCB = 0
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, True, False, False, False, False)
            treeCocau.ExpandAll()
        End If
    End Sub
End Class