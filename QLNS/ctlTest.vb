Public Class ctlTest
    Public Property labText() As String
        Get
            labText = lab.Text
        End Get
        Set(ByVal Value As String)
            lab.Text = Value
        End Set
    End Property

    'Public Property sanLbl() As String
    '    Get
    '        sanLbl = Label1.Text
    '    End Get
    '    Set(ByVal Value As String)
    '        Label1.Text = Value
    '    End Set
    'End Property
End Class
